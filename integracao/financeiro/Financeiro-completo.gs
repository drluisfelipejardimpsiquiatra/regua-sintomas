// ARQUIVO GERADO por montar-arquivo-unico.js a partir de Codigo.gs, Caixa.html, Painel.html e Estilo.html.
// Cole este arquivo inteiro no Apps Script (não precisa criar os arquivos HTML).

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
  PG: { nome: "Praia Grande", local: "Consultório Dr. Luis Felipe Jardim", unidade: 4176, pasta: "Praia Grande" },
  SO: { nome: "Sombrio", local: "Clínica Cardiovitta", unidade: 4177, pasta: "Sombrio (Cardiovitta)" },
  TO: { nome: "Torres", local: "Vida Multiclínica", unidade: 5398, pasta: "Torres (Vida Multiclínica)" },
};

const SITUACOES = ["Atendido", "Pagar depois", "Pago antes", "Cortesia", "Faltou", "Desmarcou"];
const REALIZADO = { "Atendido": 1, "Pagar depois": 1, "Pago antes": 1, "Cortesia": 1 };
// "Convênio": sessões pagas pelo convênio (ECT pela Unimed); o dinheiro vem depois, repassado pelo hospital
const FORMAS = ["Dinheiro", "Pix", "Débito", "Crédito", "Convênio"];
const CATEGORIAS_SAIDA = ["Taxa de sala · infusão", "Aluguel pago com o caixa", "Retirada pelo Dr. Luis Felipe", "Faxina", "Material de limpeza", "Material de escritório",
  "Insumos e medicação", "Estacionamento", "Lanche e copa", "Outros"];
const PAGO_COM = ["Dinheiro do caixa", "Pix", "Cartão", "Outro"];
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
  return t.evaluate()
    .setTitle(papel === "MEDICO" && !CLINICAS[c] ? "Financeiro · painel" : "Caixa · " + CLINICAS[t.clin].nome)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** URL publicada do app (/exec). Rodando pelo editor o Google devolve a /dev, que só abre para o dono:
 *  nesse caso use a propriedade URL_APP (Configurações do projeto → Propriedades do script). */
function urlApp_() {
  const p = PropertiesService.getScriptProperties().getProperty("URL_APP");
  return p || ScriptApp.getService().getUrl();
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
  let agenda = [], aviso = "";
  try { agenda = agenda_(CLINICAS[cod].unidade, data, data); }
  catch (err) { aviso = "Não consegui ler a agenda do V Saúde agora (" + err.message + "). Lance pelo botão “Paciente fora da agenda”."; }
  return {
    clinica: Object.assign({ cod: cod }, CLINICAS[cod]), data: data, hoje: hoje_(), aviso: aviso, agenda: agenda,
    lancamentos: linhas_("Lançamentos").filter(r => r.Data === data && r["Clínica"] === CLINICAS[cod].nome).map(lancOut_),
    saidas: linhas_("Saídas").filter(r => r.Data === data && r["Clínica"] === CLINICAS[cod].nome).map(saidaOut_),
    fechamento: fechamentoOut_(linhas_("Fechamentos").find(r => r.Data === data && r["Clínica"] === CLINICAS[cod].nome)),
    precos: precos_(), situacoes: SITUACOES, formas: FORMAS, categorias: CATEGORIAS_SAIDA, pagoCom: PAGO_COM,
    fotos: fotosDoDia_(cod, data).map(f => f.getName()),
  };
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
  let ags = null; try { ags = agendas_(cods.map(c => [CLINICAS[c].unidade, hoje, hoje])); } catch (e) {}
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
    const agenda = ags ? ags[3 * i] : [];
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
        status: a.status, statusTxt: statusTxt_(a.status), finalizado: !!STATUS_FINALIZADO[a.status], online: !!a.remotely };
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
  const ag = agendas_([[C.unidade, somaDiasStr_(hoje, -180), hoje], [C.unidade, hoje, "9999-12-31"]]);
  return { clinica: Object.assign({ cod: cod }, C), lista: reincidentes_(ag[0], ag[1], hoje, hora, C.nome) };
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


// ===================== TELAS EMBUTIDAS =====================

const ARQUIVOS_ = {
 "Caixa": "<!doctype html>\n<html lang=\"pt-BR\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<?!= include(\"Estilo\"); ?>\n<style>\n.it{padding:12px 14px}\n.it .nome{font-weight:800;font-size:16px}\n.it .hora{font-variant-numeric:tabular-nums;color:var(--soft);font-weight:700;margin-right:6px}\n.it .dir{text-align:right;flex-shrink:0}\n.btn.peq{padding:8px 14px;font-size:14px}\n.ed{border-top:1px solid var(--line);margin-top:10px;padding-top:6px}\n.acoes{display:flex;gap:10px;align-items:center;margin-top:10px}\n.acoes .btn{flex:1}\n.resumo{display:grid;grid-template-columns:repeat(2,1fr);gap:4px 14px;margin-top:6px;font-size:14px}\n.resumo div{display:flex;justify-content:space-between;gap:6px}\n.resumo b{font-variant-numeric:tabular-nums;white-space:nowrap}\n.big{font-size:24px;font-weight:800;color:var(--brand)}\n.sai{display:flex;justify-content:space-between;gap:8px;padding:7px 0;border-bottom:1px solid var(--line)}\n.grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px}\n.sep-foto{border-top:1px solid var(--line);margin-top:14px;padding-top:6px}\n</style>\n</head>\n<body>\n<div class=\"marca-faixa\"><svg viewBox=\"0 0 110 155\" aria-hidden=\"true\"><path d=\"M0 0H108V42L97 62 75 77 103 93 110 112V155H3V113L13 93 35 77 8 62 0 42Z\"/></svg><p class=\"nome\">Dr. Luis Felipe Jardim</p><div class=\"cargo\">Caixa do consultório</div></div>\n<div class=\"wrap\">\n  <h1 id=\"tit\">Caixa</h1>\n  <p class=\"sub-h\" id=\"subt\"></p>\n  <div class=\"tabs\"><button id=\"tDia\" class=\"on\">Dia</button><button id=\"tMes\">Mês</button><button id=\"tBusca\">Busca ativa</button></div>\n  <div id=\"vDia\">\n    <div class=\"nav\"><button id=\"ant\" aria-label=\"Dia anterior\">‹</button><input type=\"date\" id=\"data\"><button id=\"prox\" aria-label=\"Próximo dia\">›</button></div>\n    <div id=\"dia\"><div class=\"carregando\">Carregando a agenda…</div></div>\n  </div>\n  <div id=\"vMes\" hidden>\n    <div class=\"nav\"><button id=\"mAnt\" aria-label=\"Mês anterior\">‹</button><input type=\"month\" id=\"mes\"><button id=\"mProx\" aria-label=\"Próximo mês\">›</button></div>\n    <div id=\"mesBox\"></div>\n  </div>\n  <div id=\"vBusca\" hidden><div id=\"buscaBox\"></div></div>\n  <footer>Os valores vão direto para o controle financeiro do Dr. Luis Felipe.<br>Este link é só seu: não repasse.</footer>\n</div>\n<script>\nconst K = <?!= JSON.stringify(k) ?>, CLIN = <?!= JSON.stringify(clin) ?>, MEDICO = <?!= JSON.stringify(medico) ?>, URL_APP = <?!= JSON.stringify(url) ?>;\nconst PROCS = [\"Consulta\", \"Consulta 1ª vez\", \"Consulta online\", \"Infusão de escetamina\", \"Outro\"];\nlet D = null, ED = null, ocupado = false;\nconst $ = (id) => document.getElementById(id);\nconst hojeLocal = () => { const x = new Date(); x.setMinutes(x.getMinutes() - x.getTimezoneOffset()); return x.toISOString().slice(0, 10); };\n$(\"data\").value = hojeLocal(); $(\"mes\").value = hojeLocal().slice(0, 7);\nif (MEDICO) $(\"subt\").insertAdjacentHTML(\"afterend\", '<p class=\"sub-h\"><a style=\"color:var(--on-bg)\" href=\"' + URL_APP + '?k=' + encodeURIComponent(K) + '\">← voltar ao painel</a></p>');\n\n// ---------- dia ----------\nasync function carregar() {\n  ED = null; $(\"dia\").innerHTML = '<div class=\"carregando\">Carregando a agenda…</div>';\n  try { D = await call(\"dia\", K, CLIN, $(\"data\").value); render(); }\n  catch (e) { $(\"dia\").innerHTML = '<div class=\"erro\">' + esc(msgErro(e)) + \"</div>\"; }\n}\nfunction itens() {\n  const porAg = {}; D.lancamentos.forEach((l) => { if (l.agendamento) porAg[l.agendamento] = l; });\n  const lista = D.agenda.map((a) => ({ key: \"a\" + a.id, ag: a, l: porAg[a.id] || null }));\n  const naAgenda = {}; D.agenda.forEach((a) => { naAgenda[a.id] = 1; });\n  D.lancamentos.filter((l) => !l.agendamento || !naAgenda[l.agendamento]).forEach((l) => lista.push({ key: \"l\" + l.id, ag: null, l: l }));\n  return lista.sort((x, y) => ((x.ag ? x.ag.hora : x.l.hora) || \"99\").localeCompare((y.ag ? y.ag.hora : y.l.hora) || \"99\"));\n}\nfunction chipLanc(l) {\n  if (!l) return \"\";\n  if (l.situacao === \"Atendido\") {\n    const fs = D.formas.filter((f) => l.valores[f] > 0).map((f) => f === \"Crédito\" && l.parcelas > 1 ? \"Crédito \" + l.parcelas + \"x\" : f);\n    return '<span class=\"chip ' + (l.total > 0 ? \"c-ok\" : \"c-mid\") + '\">' + (l.total > 0 ? R$(l.total) + \" · \" + fs.join(\" + \") : \"Atendido sem valor\") + \"</span>\";\n  }\n  const cls = l.situacao === \"Faltou\" ? \"c-bad\" : l.situacao === \"Pagar depois\" ? \"c-mid\" : \"c-neu\";\n  return '<span class=\"chip ' + cls + '\">' + esc(l.situacao) + (l.total > 0 ? \" · \" + R$(l.total) : \"\") + \"</span>\";\n}\nfunction render() {\n  $(\"tit\").textContent = \"Caixa · \" + D.clinica.nome;\n  $(\"subt\").textContent = D.clinica.local + \" · \" + diaSemana(D.data) + \", \" + dBR(D.data);\n  const L = itens(), lancados = L.filter((x) => x.l).length;\n  const t = totais();\n  let h = D.aviso ? '<div class=\"alerta\">' + esc(D.aviso) + \"</div>\" : \"\";\n  h += '<div class=\"card\"><div class=\"row\"><div><div class=\"muted\">Recebido no dia</div><div class=\"big\">' + R$(t.entradas) + '</div></div><div class=\"muted\" style=\"text-align:right\">' +\n    lancados + \" de \" + L.length + \" lançados\" + (D.fechamento ? '<br><span class=\"chip c-ok\">Dia fechado</span>' : \"\") + \"</div></div></div>\";\n  if (!L.length) h += '<div class=\"card muted\">Nenhum paciente na agenda do V Saúde neste dia.</div>';\n  L.forEach((x) => {\n    const nome = x.ag ? x.ag.paciente : x.l.paciente, hora = x.ag ? x.ag.hora : x.l.hora, proc = x.ag ? x.ag.procedimento : x.l.procedimento;\n    const vs = x.ag ? ' · <span class=\"chip ' + (x.ag.status === 6 ? \"c-bad\" : \"c-neu\") + '\">V Saúde: ' + esc(x.ag.statusTxt.toLowerCase()) + \"</span>\" : ' · <span class=\"chip c-neu\">fora da agenda</span>';\n    h += '<div class=\"card it\" data-key=\"' + x.key + '\"><div class=\"row\"><div><div class=\"nome\"><span class=\"hora\">' + esc(hora || \"\") + \"</span>\" + esc(nome) + '</div><div class=\"muted\">' + esc(proc || \"\") + vs + \"</div></div>\" +\n      '<div class=\"dir\">' + (x.l ? chipLanc(x.l) + '<br><button class=\"lnk\" data-abrir=\"' + x.key + '\">editar</button>' : '<button class=\"btn peq\" data-abrir=\"' + x.key + '\">Lançar</button>') + \"</div></div>\" +\n      (ED && ED.key === x.key ? editor() : \"\") + \"</div>\";\n  });\n  h += '<div class=\"card it\" data-key=\"novo\">' + (ED && ED.key === \"novo\" ? \"<h2>Paciente fora da agenda</h2>\" + editor() : '<button class=\"btn2\" style=\"width:100%\" data-abrir=\"novo\">+ Paciente fora da agenda</button>') + \"</div>\";\n  h += cardSaidas() + cardFechamento(t);\n  $(\"dia\").innerHTML = h;\n}\nfunction preco(proc, forma) {\n  const P = D.precos;\n  if (forma === \"Convênio\") return P[\"ECT · repasse do hospital (Unimed)\"] || 0;\n  if (/infus|escetamin|cetamin/i.test(proc || \"\")) return P[\"Infusão de escetamina\"] || 0;\n  if (forma === \"Débito\") return P[\"Consulta · Débito\"] || 0;\n  if (forma === \"Crédito\") return P[\"Consulta · Crédito\"] || 0;\n  return P[\"Consulta · Pix ou dinheiro\"] || 0;\n}\nfunction abrir(key) {\n  if (key === \"novo\") { ED = { key: \"novo\", manual: true, paciente: \"\", procedimento: \"Consulta\", situacao: \"Atendido\", sel: {}, val: {}, parcelas: 1, obs: \"\" }; render(); return; }\n  const x = itens().find((i) => i.key === key); if (!x) return;\n  const l = x.l, a = x.ag;\n  ED = { key: key, id: l ? l.id : \"\", agendamento: a ? a.id : (l ? l.agendamento : \"\"), hora: a ? a.hora : (l ? l.hora : \"\"), manual: !a,\n    paciente: a ? a.paciente : l.paciente, procedimento: a ? a.procedimento : l.procedimento,\n    situacao: l ? l.situacao : (a && a.status === 6 ? \"Faltou\" : \"Atendido\"), sel: {}, val: {}, parcelas: (l && l.parcelas) || 1, obs: l ? l.obs : \"\" };\n  if (l) D.formas.forEach((f) => { if (l.valores[f] > 0) { ED.sel[f] = true; ED.val[f] = fmtIn(l.valores[f]); } });\n  render();\n  const el = document.querySelector('[data-key=\"' + key + '\"]'); if (el) el.scrollIntoView({ block: \"nearest\", behavior: \"smooth\" });\n}\nfunction editor() {\n  const e = ED, pagando = e.situacao === \"Atendido\";\n  let h = '<div class=\"ed\">';\n  if (e.manual) {\n    h += '<label class=\"l\">Paciente</label><input class=\"f\" id=\"eNome\" value=\"' + esc(e.paciente) + '\" placeholder=\"Nome completo\">';\n    h += '<label class=\"l\">Procedimento</label><select class=\"f\" id=\"eProc\">' + PROCS.map((p) => \"<option\" + (p === e.procedimento ? \" selected\" : \"\") + \">\" + p + \"</option>\").join(\"\") + \"</select>\";\n  }\n  h += '<label class=\"l\">Situação</label><div class=\"opts\">' + D.situacoes.map((s) => '<button type=\"button\" data-sit=\"' + s + '\" class=\"' + (s === e.situacao ? \"on\" : \"\") + '\">' + s + \"</button>\").join(\"\") + \"</div>\";\n  if (pagando) {\n    const desc = D.precos[\"Consulta com desconto\"];\n    h += '<label class=\"l\">Forma de pagamento (pode marcar mais de uma)</label><div class=\"opts\">' + D.formas.map((f) => '<button type=\"button\" data-forma=\"' + f + '\" class=\"' + (e.sel[f] ? \"on\" : \"\") + '\">' + f + \"</button>\").join(\"\") + \"</div>\";\n    D.formas.filter((f) => e.sel[f]).forEach((f) => {\n      h += '<div class=\"pg\"><span class=\"nm\">' + f + '</span><input class=\"f\" inputmode=\"decimal\" data-val=\"' + f + '\" value=\"' + esc(e.val[f] || \"\") + '\" placeholder=\"R$\">' +\n        (f === \"Crédito\" ? '<select class=\"f\" id=\"eParc\">' + Array.from({ length: 12 }, (_, i) => '<option value=\"' + (i + 1) + '\"' + (+e.parcelas === i + 1 ? \" selected\" : \"\") + \">\" + (i + 1) + \"x</option>\").join(\"\") + \"</select>\" : \"<span></span>\") + \"</div>\";\n    });\n    if (desc && !/infus|escetamin|cetamin/i.test(e.procedimento || \"\")) h += '<button type=\"button\" class=\"lnk\" id=\"eDesc\" style=\"color:var(--brand)\">Usar valor com desconto (' + R$(desc) + \")</button>\";\n  } else if (e.situacao === \"Pagar depois\") h += '<p class=\"muted\">Fica como “a receber” no controle. Quando o paciente pagar, volte aqui, marque “Atendido” e informe o valor.</p>';\n  else if (e.situacao === \"Pago antes\") h += '<p class=\"muted\">Use só se o valor JÁ foi lançado no dia em que o paciente pagou. Se ainda não foi lançado em lugar nenhum, marque “Atendido” com o valor e escreva na observação “pago em dd/mm”.</p>';\n  h += '<label class=\"l\">Observação (opcional)</label><textarea class=\"f\" id=\"eObs\" maxlength=\"300\">' + esc(e.obs || \"\") + \"</textarea>\";\n  h += '<div class=\"acoes\"><button class=\"btn\" id=\"eSalvar\">Salvar</button><button class=\"lnk\" id=\"eCanc\">cancelar</button>' + (e.id ? '<button class=\"lnk\" id=\"eExc\">excluir</button>' : \"\") + \"</div></div>\";\n  return h;\n}\nfunction lerEditor() {\n  if (!ED) return;\n  if ($(\"eNome\")) ED.paciente = $(\"eNome\").value;\n  if ($(\"eProc\")) ED.procedimento = $(\"eProc\").value;\n  if ($(\"eObs\")) ED.obs = $(\"eObs\").value;\n  if ($(\"eParc\")) ED.parcelas = +$(\"eParc\").value;\n  document.querySelectorAll(\"[data-val]\").forEach((i) => { ED.val[i.dataset.val] = i.value; });\n}\nasync function salvar() {\n  lerEditor();\n  const e = ED, o = { id: e.id, data: D.data, agendamento: e.agendamento || \"\", hora: e.hora || \"\", paciente: e.paciente, procedimento: e.procedimento, situacao: e.situacao, parcelas: e.parcelas, obs: e.obs };\n  let total = 0;\n  if (e.situacao === \"Atendido\") D.formas.forEach((f) => { o[f] = e.sel[f] ? num(e.val[f]) : 0; total += o[f]; });\n  if (e.situacao === \"Atendido\" && !total && !confirm(\"Atendido sem valor nenhum. Salvar assim mesmo?\")) return;\n  await enviar(\"salvarLancamento\", o);\n}\nasync function enviar(fn, ...args) {\n  if (ocupado) return; ocupado = true;\n  document.querySelectorAll(\"#dia .btn\").forEach((b) => (b.disabled = true));\n  try { await call(fn, K, CLIN, ...args); await carregar(); }\n  catch (e) { alert(msgErro(e)); document.querySelectorAll(\"#dia .btn\").forEach((b) => (b.disabled = false)); }\n  finally { ocupado = false; }\n}\n\n// ---------- saídas e fechamento ----------\nfunction totais() {\n  const f = {}; D.formas.forEach((x) => (f[x] = 0));\n  D.lancamentos.forEach((l) => D.formas.forEach((x) => (f[x] += l.valores[x])));\n  const entradas = D.formas.reduce((s, x) => s + f[x], 0), saidas = D.saidas.reduce((s, x) => s + x.valor, 0);\n  const saidasDin = D.saidas.filter((x) => x.pagoCom === \"Dinheiro do caixa\").reduce((s, x) => s + x.valor, 0);\n  return { formas: f, entradas, saidas, saidasDin, esperado: f[\"Dinheiro\"] - saidasDin };\n}\nfunction cardSaidas() {\n  let h = '<div class=\"card\"><h2>Saídas do dia</h2>';\n  if (!D.saidas.length) h += '<p class=\"muted\">Nenhuma saída lançada. Ex.: taxa de sala da infusão, faxina, material.</p>';\n  D.saidas.forEach((s) => { h += '<div class=\"sai\"><div><b>' + esc(s.categoria) + \"</b>\" + (s.descricao ? \" · \" + esc(s.descricao) : \"\") + '<div class=\"muted\">' + esc(s.pagoCom) + '</div></div><div style=\"text-align:right\"><b>' + R$(s.valor) + '</b><br><button class=\"lnk\" data-excsai=\"' + esc(s.id) + '\">excluir</button></div></div>'; });\n  h += '<div class=\"grid2\" style=\"margin-top:8px\"><div><label class=\"l\">Tipo</label><select class=\"f\" id=\"sCat\">' + D.categorias.map((c) => \"<option>\" + c + \"</option>\").join(\"\") + \"</select></div>\" +\n    '<div><label class=\"l\">Valor</label><input class=\"f\" id=\"sVal\" inputmode=\"decimal\" placeholder=\"R$\"></div></div>' +\n    '<label class=\"l\">Descrição (opcional)</label><input class=\"f\" id=\"sDesc\" maxlength=\"200\" placeholder=\"Ex.: 3 infusões no hospital\">' +\n    '<label class=\"l\">Pago com</label><select class=\"f\" id=\"sPago\">' + D.pagoCom.map((c) => \"<option>\" + c + \"</option>\").join(\"\") + \"</select>\" +\n    '<div class=\"acoes\"><button class=\"btn\" id=\"sAdd\">Adicionar saída</button></div></div>';\n  return h;\n}\nfunction cardFechamento(t) {\n  const F = D.fechamento;\n  let h = '<div class=\"card\"><h2>Fechamento do dia</h2><div class=\"resumo\">' + D.formas.map((f) => \"<div>\" + f + \"<b>\" + R$(t.formas[f]) + \"</b></div>\").join(\"\") +\n    \"<div><strong>Total recebido</strong><b>\" + R$(t.entradas) + \"</b></div><div>Saídas<b>\" + R$(t.saidas) + \"</b></div></div>\";\n  h += '<p class=\"muted\" style=\"margin:10px 0 0\">Dinheiro que deveria estar na gaveta (dinheiro recebido − saídas pagas com o dinheiro do caixa): <b style=\"color:var(--ink)\">' + R$(t.esperado) + \"</b></p>\";\n  h += '<div class=\"grid2\"><div><label class=\"l\">Dinheiro contado na gaveta</label><input class=\"f\" id=\"fCont\" inputmode=\"decimal\" value=\"' + (F && F.contado != null ? fmtIn(F.contado) : \"\") + '\" placeholder=\"R$\"></div>' +\n    '<div><label class=\"l\">Observação</label><input class=\"f\" id=\"fObs\" maxlength=\"300\" value=\"' + esc(F ? F.obs : \"\") + '\"></div></div>';\n  if (F) h += '<p class=\"muted\">Fechado em ' + esc(F.fechadoEm) + (F.diferenca != null ? ' · diferença <span class=\"chip ' + (Math.abs(F.diferenca) < 0.01 ? \"c-ok\" : \"c-bad\") + '\">' + R$(F.diferenca) + \"</span>\" : \"\") + \". Se mudar algum lançamento, feche de novo.</p>\";\n  h += '<div class=\"acoes\"><button class=\"btn\" id=\"fFechar\">' + (F ? \"Fechar de novo\" : \"Fechar o dia\") + \"</button></div>\";\n  h += '<div class=\"sep-foto\"><label class=\"l\">Foto do fechamento (vai para a pasta da clínica no Drive do Dr. Luis Felipe)</label>' +\n    (D.fotos && D.fotos.length ? '<p class=\"muted\">Enviada' + (D.fotos.length > 1 ? \"s \" + D.fotos.length + \" fotos\" : \"\") + ' ✓</p>' : \"\") +\n    '<input type=\"file\" id=\"fFoto\" accept=\"image/*\" hidden><button class=\"btn2\" id=\"fFotoBtn\" style=\"width:100%\">' + (D.fotos && D.fotos.length ? \"Enviar outra foto\" : \"📷 Enviar foto do fechamento\") + \"</button></div></div>\";\n  return h;\n}\n\n// ---------- mês ----------\nasync function carregarMes() {\n  $(\"mesBox\").innerHTML = '<div class=\"carregando\">Carregando…</div>';\n  try {\n    const M = await call(\"mesClinica\", K, CLIN, $(\"mes\").value);\n    let h = '<div class=\"card\"><h2>' + esc(mesBR(M.mes)) + '</h2><div class=\"resumo\">' + Object.keys(M.formas).map((f) => \"<div>\" + f + \"<b>\" + R$(M.formas[f]) + \"</b></div>\").join(\"\") +\n      \"<div><strong>Total recebido</strong><b>\" + R$(M.entradas) + \"</b></div><div>Saídas<b>\" + R$(M.saidas) + \"</b></div><div>Atendimentos<b>\" + M.atendimentos + \"</b></div><div>Faltas<b>\" + M.faltas + \"</b></div></div></div>\";\n    h += '<div class=\"card\"><table><tr><th>Dia</th><th class=\"n\">Pac.</th><th class=\"n\">Recebido</th><th class=\"n\">Saídas</th><th>Caixa</th></tr>' +\n      (M.dias.length ? M.dias.map((d) => '<tr><td><button class=\"lnk\" style=\"color:var(--ink)\" data-irdia=\"' + d.data + '\">' + dBR(d.data).slice(0, 5) + \"</button> \" + diaSemana(d.data).slice(0, 3) + '</td><td class=\"n\">' + d.atendimentos + '</td><td class=\"n\">' + R$(d.entradas) + '</td><td class=\"n\">' + R$(d.saidas) + \"</td><td>\" +\n        (d.fechado ? '<span class=\"chip ' + (d.diferenca && Math.abs(d.diferenca) >= 0.01 ? \"c-bad\" : \"c-ok\") + '\">' + (d.diferenca && Math.abs(d.diferenca) >= 0.01 ? R$(d.diferenca) : \"fechado\") + \"</span>\" : '<span class=\"chip c-mid\">aberto</span>') + \"</td></tr>\").join(\"\")\n        : '<tr><td colspan=\"5\" class=\"muted\">Nada lançado neste mês.</td></tr>') + \"</table></div>\";\n    $(\"mesBox\").innerHTML = h;\n  } catch (e) { $(\"mesBox\").innerHTML = '<div class=\"erro\">' + esc(msgErro(e)) + \"</div>\"; }\n}\n\n// ---------- busca ativa ----------\nasync function carregarBusca() {\n  $(\"buscaBox\").innerHTML = '<div class=\"carregando\">Consultando a agenda dos últimos 6 meses…</div>';\n  try {\n    const B = await call(\"buscaAtiva\", K, CLIN);\n    $(\"buscaBox\").innerHTML = '<div class=\"card\"><h2>Pacientes que faltaram 2 ou mais vezes</h2><p class=\"muted\">Últimos 6 meses, pela agenda do V Saúde. Comece por quem faltou nas 2 últimas e por quem está <b>sem retorno marcado</b>: o botão abre o WhatsApp com uma mensagem pronta. Para quem já tem a próxima consulta marcada, confirme a presença na véspera.</p>' +\n      htmlBuscaAtiva(B.lista, false) + \"</div>\";\n  } catch (e) { $(\"buscaBox\").innerHTML = '<div class=\"erro\">' + esc(msgErro(e)) + \"</div>\"; }\n}\n\n// ---------- eventos ----------\ndocument.addEventListener(\"click\", (ev) => {\n  const t = ev.target.closest(\"button\"); if (!t) return;\n  if (t.dataset.abrir) { abrir(t.dataset.abrir); return; }\n  if (t.dataset.sit) { lerEditor(); ED.situacao = t.dataset.sit; render(); return; }\n  if (t.dataset.forma) {\n    lerEditor(); const f = t.dataset.forma; ED.sel[f] = !ED.sel[f];\n    const marcadas = D.formas.filter((x) => ED.sel[x]);\n    if (ED.sel[f] && marcadas.length === 1 && !ED.val[f]) ED.val[f] = fmtIn(preco(ED.procedimento, f));\n    if (!ED.sel[f]) delete ED.val[f];\n    render(); const i = document.querySelector('[data-val=\"' + f + '\"]'); if (i && marcadas.length > 1) i.focus(); return;\n  }\n  if (t.dataset.excsai) { if (confirm(\"Excluir esta saída?\")) enviar(\"excluirSaida\", t.dataset.excsai); return; }\n  if (t.dataset.irdia) { $(\"data\").value = t.dataset.irdia; trocaAba(\"dia\"); carregar(); return; }\n  switch (t.id) {\n    case \"eSalvar\": salvar(); break;\n    case \"eDesc\": { lerEditor(); const f = D.formas.find((x) => ED.sel[x]) || \"Pix\"; D.formas.forEach((x) => { if (x !== f) { delete ED.sel[x]; delete ED.val[x]; } });\n      ED.sel[f] = true; ED.val[f] = fmtIn(D.precos[\"Consulta com desconto\"]); ED.obs = ED.obs || \"Desconto\"; render(); break; }\n    case \"eCanc\": ED = null; render(); break;\n    case \"eExc\": if (confirm(\"Excluir este lançamento?\")) enviar(\"excluirLancamento\", ED.id); break;\n    case \"sAdd\": enviar(\"salvarSaida\", { data: D.data, categoria: $(\"sCat\").value, valor: num($(\"sVal\").value), descricao: $(\"sDesc\").value, pagoCom: $(\"sPago\").value }); break;\n    case \"fFotoBtn\": $(\"fFoto\").click(); break;\n    case \"fFechar\": enviar(\"fecharDia\", D.data, $(\"fCont\").value.trim() === \"\" ? \"\" : num($(\"fCont\").value), $(\"fObs\").value); break;\n    case \"ant\": $(\"data\").value = somaDias($(\"data\").value, -1); carregar(); break;\n    case \"prox\": $(\"data\").value = somaDias($(\"data\").value, 1); carregar(); break;\n    case \"mAnt\": $(\"mes\").value = somaMes($(\"mes\").value, -1); carregarMes(); break;\n    case \"mProx\": $(\"mes\").value = somaMes($(\"mes\").value, 1); carregarMes(); break;\n    case \"tDia\": trocaAba(\"dia\"); break;\n    case \"tMes\": trocaAba(\"mes\"); carregarMes(); break;\n    case \"tBusca\": trocaAba(\"busca\"); carregarBusca(); break;\n  }\n});\n// reduz a foto no próprio celular (até 1600 px) antes de enviar\nfunction reduzir(arquivo) {\n  return new Promise((ok, no) => {\n    const img = new Image(), url = URL.createObjectURL(arquivo);\n    img.onload = () => { const r = Math.min(1, 1600 / Math.max(img.width, img.height)), cv = document.createElement(\"canvas\");\n      cv.width = Math.round(img.width * r); cv.height = Math.round(img.height * r); cv.getContext(\"2d\").drawImage(img, 0, 0, cv.width, cv.height);\n      URL.revokeObjectURL(url); ok(cv.toDataURL(\"image/jpeg\", 0.85)); };\n    img.onerror = () => no(new Error(\"Não consegui abrir a foto.\"));\n    img.src = url;\n  });\n}\ndocument.addEventListener(\"change\", async (ev) => {\n  if (ev.target.id === \"fFoto\" && ev.target.files[0]) {\n    const b = $(\"fFotoBtn\"); b.disabled = true; b.textContent = \"Enviando a foto…\";\n    try { await call(\"enviarFoto\", K, CLIN, D.data, await reduzir(ev.target.files[0])); await carregar(); }\n    catch (e) { alert(msgErro(e)); b.disabled = false; b.textContent = \"📷 Enviar foto do fechamento\"; }\n    return;\n  }\n  if (ev.target.id === \"data\" && ev.target.value) carregar();\n  if (ev.target.id === \"mes\" && ev.target.value) carregarMes();\n  if (ev.target.id === \"eProc\") lerEditor();\n});\nfunction trocaAba(a) {\n  [\"dia\", \"mes\", \"busca\"].forEach((x) => { const id = x[0].toUpperCase() + x.slice(1); $(\"v\" + id).hidden = a !== x; $(\"t\" + id).classList.toggle(\"on\", a === x); });\n}\ncarregar();\n</script>\n</body>\n</html>\n",
 "Painel": "<!doctype html>\n<html lang=\"pt-BR\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<?!= include(\"Estilo\"); ?>\n<style>\n.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:12px 0}\n.kpi{background:var(--card);border-radius:16px;padding:12px 14px;box-shadow:var(--shadow-card)}\n.kpi .v{font-size:22px;font-weight:800;font-variant-numeric:tabular-nums;margin-top:2px}\n.kpi .s{font-size:12px;color:var(--soft);margin-top:2px}\n.clins{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}\n.clins .card{margin:0}\n.clins h2{margin-bottom:2px}\n.linha{display:flex;justify-content:space-between;gap:8px;padding:4px 0;font-size:14px}\n.linha span:last-child{font-variant-numeric:tabular-nums;white-space:nowrap}\n.linha.ind{padding-left:12px;color:var(--soft);font-size:13px}\n.linha.tot{border-top:1px solid var(--line);margin-top:6px;padding-top:8px;font-weight:800;font-size:16px}\n.sep{height:1px;background:var(--line);margin:6px 0}\n.neg{color:var(--bad-text)}\n.barra{height:8px;border-radius:4px;background:var(--brand-soft);overflow:hidden;display:flex;margin:6px 0 2px}\n.barra i{display:block;height:100%}\n.leg{display:flex;flex-wrap:wrap;gap:4px 12px;font-size:12px;color:var(--soft)}\n.leg i{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:4px;vertical-align:-1px}\n.topo{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap}\n.topo>div:first-child{flex:1 1 360px}\n.ano{background:var(--card);border-radius:16px;padding:12px 14px 10px;box-shadow:var(--shadow-card);flex:0 1 330px;min-width:260px;margin-top:10px}\n.ano .v{font-size:24px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.1}\n.ano .s{font-size:12px;color:var(--soft);margin-top:2px;min-height:16px}\n.barras{display:grid;grid-template-columns:repeat(12,1fr);gap:2px;align-items:end;height:56px;margin-top:8px}\n.barras button{all:unset;cursor:pointer;height:100%;display:flex;align-items:flex-end;border-radius:4px}\n.barras button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}\n.barras i{display:block;width:100%;background:var(--brand);border-radius:4px 4px 0 0;min-height:0}\n.barras button.on i{background:var(--accent)}\n.barras i{display:flex;flex-direction:column;overflow:hidden}\n.barras i b,.leg i.est,.barra i.est,.estim{background-image:repeating-linear-gradient(135deg,rgba(251,248,238,.6) 0 2px,transparent 2px 5px)}\n.barras i b{display:block;width:100%}\n.estim{display:inline-block;width:10px;height:10px;border-radius:2px;background-color:var(--brand);vertical-align:-1px;margin-right:4px}\n.barras button.vazio i{background:var(--line);height:2px!important}\n.meses{display:grid;grid-template-columns:repeat(12,1fr);gap:2px;font-size:10px;color:var(--soft);text-align:center;margin-top:3px}\n@media (max-width:860px){.clins{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}}\n</style>\n</head>\n<body>\n<div class=\"marca-faixa\"><svg viewBox=\"0 0 110 155\" aria-hidden=\"true\"><path d=\"M0 0H108V42L97 62 75 77 103 93 110 112V155H3V113L13 93 35 77 8 62 0 42Z\"/></svg><p class=\"nome\">Dr. Luis Felipe Jardim</p><div class=\"cargo\">Financeiro do consultório</div></div>\n<div class=\"wrap largo\">\n  <div class=\"topo\">\n    <div><h1>Painel financeiro</h1>\n    <p class=\"sub-h\">Praia Grande, Sombrio (Cardiovitta) e Torres (Vida Multiclínica) somados. Só este link vê tudo.</p></div>\n    <div class=\"ano\" id=\"ano\" hidden></div>\n  </div>\n  <div class=\"nav\" style=\"max-width:420px\"><button id=\"mAnt\" aria-label=\"Mês anterior\">‹</button><input type=\"month\" id=\"mes\"><button id=\"mProx\" aria-label=\"Próximo mês\">›</button></div>\n  <div id=\"box\"><div class=\"carregando\">Calculando o mês…</div></div>\n  <footer>Custos fixos e preços: abas “Custos fixos” e “Preços” da planilha. “Convênio” = sessões pela Unimed, repassadas depois pelo hospital.</footer>\n</div>\n<script>\nconst K = <?!= JSON.stringify(k) ?>, URL_APP = <?!= JSON.stringify(url) ?>;\nconst $ = (id) => document.getElementById(id);\nconst CORES = { \"Pix\": \"#38584B\", \"Dinheiro\": \"#88764F\", \"Débito\": \"#E89925\", \"Crédito\": \"#C3622D\", \"Convênio\": \"#A3B2A4\", \"Estimado pela agenda\": \"#9A946A\" };\nconst hojeLocal = () => { const x = new Date(); x.setMinutes(x.getMinutes() - x.getTimezoneOffset()); return x.toISOString().slice(0, 10); };\n$(\"mes\").value = hojeLocal().slice(0, 7);\nconst linha = (a, b, cls) => '<div class=\"linha ' + (cls || \"\") + '\"><span>' + a + \"</span><span>\" + b + \"</span></div>\";\nconst barra = (formas, total) => '<div class=\"barra\">' + Object.keys(formas).map((f) => formas[f] > 0 ? '<i style=\"width:' + (formas[f] / total * 100) + \"%;background-color:\" + CORES[f] + '\" ' + (f === \"Estimado pela agenda\" ? 'class=\"est\" ' : \"\") + 'title=\"' + f + '\"></i>' : \"\").join(\"\") + \"</div>\" +\n  '<div class=\"leg\">' + Object.keys(formas).filter((f) => formas[f] > 0).map((f) => \"<span><i \" + (f === \"Estimado pela agenda\" ? 'class=\"est\" ' : \"\") + \"style=\\\"background-color:\" + CORES[f] + \"\\\"></i>\" + f + \" \" + Math.round(formas[f] / total * 100) + \"%</span>\").join(\"\") + \"</div>\";\n\nasync function carregar() {\n  $(\"box\").innerHTML = '<div class=\"carregando\">Calculando o mês…</div>';\n  try { render(await call(\"painel\", K, $(\"mes\").value)); }\n  catch (e) { $(\"box\").innerHTML = '<div class=\"erro\">' + esc(msgErro(e)) + \"</div>\"; }\n}\nconst MES3 = [\"jan\", \"fev\", \"mar\", \"abr\", \"mai\", \"jun\", \"jul\", \"ago\", \"set\", \"out\", \"nov\", \"dez\"];\nfunction renderAno(A, mesAtual) {\n  if (!A) return;\n  const max = Math.max(1, ...A.meses.map((m) => m.total));\n  const legenda = \"média \" + R$(A.media) + \" por mês · \" + A.nMeses + (A.nMeses === 1 ? \" mês\" : \" meses\") + \" com entradas\" +\n    (A.estimado ? \" · inclui \" + R$(A.estimado) + \" estimado pela agenda\" : \"\");\n  $(\"ano\").hidden = false;\n  $(\"ano\").innerHTML = '<div class=\"muted\">Acumulado de ' + A.ano + '</div><div class=\"v\">' + R$(A.total) + '</div><div class=\"s\" id=\"anoS\">' + legenda + \"</div>\" +\n    '<div class=\"barras\" role=\"list\">' + A.meses.map((m, i) => '<button type=\"button\" role=\"listitem\" data-mesano=\"' + m.mes + '\" class=\"' + (m.mes === mesAtual ? \"on \" : \"\") + (m.total ? \"\" : \"vazio\") +\n      '\" aria-label=\"' + MESES[i] + \": \" + R$(m.total) + '\" data-dica=\"' + MES3[i] + \": \" + R$(m.total) + (m.estimado ? \" · \" + R$(m.estimado) + \" estimado\" : \"\") + '\"><i style=\"height:' + (m.total / max * 100).toFixed(1) + '%\">' +\n      (m.estimado ? '<b style=\"height:' + (m.estimado / m.total * 100).toFixed(1) + '%\"></b>' : \"\") + \"</i></button>\").join(\"\") + \"</div>\" +\n    '<div class=\"meses\" aria-hidden=\"true\">' + MES3.map((x) => \"<span>\" + x[0].toUpperCase() + \"</span>\").join(\"\") + \"</div>\";\n  const S = $(\"anoS\");\n  $(\"ano\").querySelectorAll(\"[data-dica]\").forEach((b) => {\n    const mostra = () => { S.textContent = b.dataset.dica; }, volta = () => { S.textContent = legenda; };\n    b.addEventListener(\"mouseenter\", mostra); b.addEventListener(\"focus\", mostra);\n    b.addEventListener(\"mouseleave\", volta); b.addEventListener(\"blur\", volta);\n  });\n}\nfunction render(P) {\n  renderAno(P.ano, P.mes);\n  const T = P.total, margem = T.entradas ? Math.round(T.resultado / T.entradas * 100) : 0;\n  let h = '<div class=\"kpis\">' +\n    '<div class=\"kpi\"><div class=\"muted\">Entradas</div><div class=\"v\">' + R$(T.entradas) + '</div><div class=\"s\">' + esc(mesBR(P.mes)) + (T.estimado ? ' · <span class=\"estim\"></span>inclui ' + R$(T.estimado) + \" estimado pela agenda\" : \"\") + \"</div></div>\" +\n    '<div class=\"kpi\"><div class=\"muted\">Custos e saídas</div><div class=\"v\">' + R$(T.custos) + '</div><div class=\"s\">fixos + taxa de infusão + caixa</div></div>' +\n    '<div class=\"kpi\"><div class=\"muted\">Resultado</div><div class=\"v' + (T.resultado < 0 ? \" neg\" : \"\") + '\">' + R$(T.resultado) + '</div><div class=\"s\">margem ' + margem + \"%</div></div>\" +\n    '<div class=\"kpi\"><div class=\"muted\">Atendimentos</div><div class=\"v\">' + T.atendimentos + '</div><div class=\"s\">' + T.infusoes + \" infusões · ticket médio \" + R$(T.atendimentos ? T.entradas / T.atendimentos : 0) + \"</div></div></div>\";\n  if (T.entradas) h += '<div class=\"card\">' + linha(\"<b>Por forma de pagamento</b>\", \"\") + barra(T.formas, T.entradas) + \"</div>\";\n\n  h += '<div class=\"clins\">';\n  P.clinicas.forEach((c) => {\n    h += '<div class=\"card\"><h2>' + esc(c.nome) + '</h2><div class=\"muted\">' + esc(c.local) + \"</div>\";\n    h += linha(\"<b>Entradas</b>\", \"<b>\" + R$(c.entradas) + \"</b>\");\n    Object.keys(c.formas).forEach((f) => { if (c.formas[f]) h += linha(f === \"Estimado pela agenda\" ? '<span class=\"estim\"></span>Estimado pela agenda · ' + c.nEstimado + \" atend. sem caixa\" : f, R$(c.formas[f]), \"ind\"); });\n    if (c.entradas) h += barra(c.formas, c.entradas);\n    h += linha(\"Atendimentos\", c.atendimentos + (c.infusoes ? \" (\" + c.infusoes + \" infusões)\" : \"\")) +\n      (c.faltas ? linha(\"Faltas lançadas\", c.faltas, \"ind\") : \"\") + (c.aReceber ? linha(\"A receber (pagar depois)\", c.aReceber, \"ind\") : \"\");\n    h += '<div class=\"sep\"></div>' + linha(\"<b>Custos</b>\", \"<b>\" + R$(c.custos) + \"</b>\");\n    c.fixos.forEach((f) => { h += linha(esc(f.descricao), f.valor == null ? '<span class=\"chip c-mid\">preencher</span>' : R$(f.valor), \"ind\"); });\n    if (c.aluguelPagoCaixa) h += linha(\"↳ aluguel já pago com o caixa\", R$(c.aluguelPagoCaixa), \"ind\");\n    if (c.infusoes || c.taxaPagaCaixa) h += linha(\"Taxa de sala · \" + c.infusoes + \" × \" + R$(c.taxaPorInfusao), R$(c.taxaDevida), \"ind\") +\n      (c.taxaPagaCaixa ? linha(\"↳ já paga pelo caixa\", R$(c.taxaPagaCaixa), \"ind\") : \"\");\n    if (c.saidasOutras) Object.keys(c.saidasPorCategoria).filter((k) => [\"Taxa de sala · infusão\", \"Aluguel pago com o caixa\", \"Retirada pelo Dr. Luis Felipe\"].indexOf(k) < 0).forEach((k) => { h += linha(\"Caixa · \" + esc(k), R$(c.saidasPorCategoria[k]), \"ind\"); });\n    h += linha(\"Resultado\", '<span class=\"' + (c.resultado < 0 ? \"neg\" : \"\") + '\">' + R$(c.resultado) + \"</span>\", \"tot\");\n    if (c.retiradas) h += linha(\"Retirado do caixa por você\", R$(c.retiradas), \"ind\");\n    h += '<a class=\"lnk\" style=\"display:inline-block;margin-top:8px\"  href=\"' + URL_APP + '?k=' + encodeURIComponent(K) + \"&c=\" + c.cod + '\">Abrir o caixa desta clínica →</a></div>';\n  });\n  h += \"</div>\";\n\n  // faltas pela agenda do V Saúde\n  const tx = (f) => f.atendidos + f.faltas ? Math.round(f.faltas / (f.atendidos + f.faltas) * 100) + \"%\" : \"—\";\n  h += '<div class=\"card\"><h2>Faltas no mês (agenda do V Saúde)</h2><div class=\"rol\"><table><tr><th>Clínica</th><th class=\"n\">Atendidos</th><th class=\"n\">Faltas</th><th class=\"n\">Taxa</th><th class=\"n\">Acompanhamento</th><th class=\"n\">1ª vez</th><th class=\"n\">Infusão</th><th class=\"n\">Receita potencial perdida</th></tr>';\n  P.clinicas.forEach((c) => { const F = c.faltasVS, t = (n) => { const x = F.porTipo.find((y) => y.tipo === n); return x ? tx(x) + ' <span class=\"muted\">(' + x.faltas + \")</span>\" : \"—\"; };\n    h += \"<tr><td>\" + esc(c.nome) + '</td><td class=\"n\">' + F.atendidos + '</td><td class=\"n\">' + F.faltas + '</td><td class=\"n\"><b>' + tx(F) + '</b></td><td class=\"n\">' + t(\"Acompanhamento\") + '</td><td class=\"n\">' + t(\"1ª vez\") + '</td><td class=\"n\">' + t(\"Infusão\") + '</td><td class=\"n\">' + R$(F.perdida) + \"</td></tr>\"; });\n  h += '<tr class=\"tot\"><td>Total</td><td class=\"n\">' + T.atendidosVS + '</td><td class=\"n\">' + T.faltasVS + '</td><td class=\"n\">' + tx({ atendidos: T.atendidosVS, faltas: T.faltasVS }) + '</td><td></td><td></td><td></td><td class=\"n\">' + R$(T.perdida) + \"</td></tr></table></div>\" +\n    '<p class=\"muted\">Conta o que o V Saúde marcou como “não compareceu” até agora. Receita perdida = faltas × preço de tabela (consulta Pix ou infusão); é potencial, parte dos horários vira encaixe.</p></div>';\n  h += '<div class=\"card\"><h2>Busca ativa (' + P.buscaAtiva.length + ')</h2><p class=\"muted\">Pacientes com 2 ou mais faltas nos últimos 6 meses. Primeiro quem faltou nas 2 últimas e quem está sem retorno marcado. Cada secretária vê a lista da própria clínica.</p>' + htmlBuscaAtiva(P.buscaAtiva, true) + \"</div>\";\n  h += '<div class=\"card\"><h2>Pendências (' + P.pendencias.length + ')</h2>';\n  h += P.pendencias.length ? '<div class=\"rol\"><table><tr><th>Clínica</th><th>Data</th><th>Paciente</th><th>O que falta</th></tr>' +\n    P.pendencias.map((p) => \"<tr><td>\" + esc(p.clinica) + \"</td><td>\" + dBR(p.data).slice(0, 5) + (p.hora ? \" \" + esc(p.hora) : \"\") + \"</td><td>\" + esc(p.paciente) + (p.procedimento ? '<div class=\"muted\">' + esc(p.procedimento) + \"</div>\" : \"\") + \"</td><td>\" + esc(p.motivo) + \"</td></tr>\").join(\"\") + \"</table></div>\"\n    : '<p class=\"muted\">Nada pendente: tudo o que estava na agenda até hoje foi lançado.</p>';\n  h += '<p class=\"muted\">Conferência da agenda do V Saúde a partir de ' + dBR(P.inicio) + ', início do sistema.</p>';\n  h += \"</div>\";\n\n  const dias = {}; P.clinicas.forEach((c) => c.dias.forEach((d) => { (dias[d.data] = dias[d.data] || {})[c.nome] = d; }));\n  const ds = Object.keys(dias).sort();\n  h += '<div class=\"card\"><h2>Dia a dia</h2><div class=\"rol\"><table><tr><th>Dia</th>' + P.clinicas.map((c) => '<th class=\"n\">' + esc(c.nome) + \"</th>\").join(\"\") + '<th class=\"n\">Total</th></tr>';\n  ds.forEach((d) => {\n    let t = 0; h += \"<tr><td>\" + dBR(d).slice(0, 5) + \" \" + diaSemana(d).slice(0, 3) + \"</td>\";\n    P.clinicas.forEach((c) => { const x = dias[d][c.nome]; t += x ? x.entradas : 0; h += '<td class=\"n\">' + (x ? (x.estimado ? '<span class=\"estim\" title=\"inclui ' + R$(x.estimado) + ' estimado pela agenda\"></span>' : \"\") + R$(x.entradas) + (x.fechado || x.estimado === x.entradas ? \"\" : ' <span class=\"chip c-mid\" title=\"caixa não fechado\">!</span>') : \"—\") + \"</td>\"; });\n    h += '<td class=\"n\"><b>' + R$(t) + \"</b></td></tr>\";\n  });\n  if (!ds.length) h += '<tr><td colspan=\"5\" class=\"muted\">Nada lançado neste mês.</td></tr>';\n  else h += '<tr class=\"tot\"><td>Total</td>' + P.clinicas.map((c) => '<td class=\"n\">' + R$(c.entradas) + \"</td>\").join(\"\") + '<td class=\"n\">' + R$(T.entradas) + \"</td></tr>\";\n  h += '</table></div>' + (T.estimado ? '<p class=\"muted\"><span class=\"estim\"></span>Dia sem caixa da secretária: valor estimado pelas consultas finalizadas no V Saúde × preço de tabela (consulta R$ 600, infusão R$ 850; ECT, retorno e cortesia não entram). Quando o caixa real do dia for lançado, a estimativa sai sozinha.</p>' : \"\") + '<p class=\"muted\"><a style=\"color:var(--brand)\" href=\"' + esc(P.planilha) + '\" target=\"_blank\" rel=\"noopener\">Abrir a planilha completa</a></p></div>';\n  $(\"box\").innerHTML = h;\n}\ndocument.addEventListener(\"click\", (ev) => {\n  const t = ev.target.closest(\"button\"); if (!t) return;\n  if (t.dataset.mesano) { $(\"mes\").value = t.dataset.mesano; carregar(); return; }\n  if (t.id === \"mAnt\") { $(\"mes\").value = somaMes($(\"mes\").value, -1); carregar(); }\n  if (t.id === \"mProx\") { $(\"mes\").value = somaMes($(\"mes\").value, 1); carregar(); }\n});\n$(\"mes\").addEventListener(\"change\", () => { if ($(\"mes\").value) carregar(); });\ncarregar();\n</script>\n</body>\n</html>\n",
 "Estilo": "<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link href=\"https://fonts.googleapis.com/css2?family=Figtree:wght@400;500;600;700;800&family=Marcellus&display=swap\" rel=\"stylesheet\">\n<style>\n/* Marca Dr. Luis Felipe Jardim · tema Oliva (mesmos tokens de marca/marca.css) */\n:root{\n  --oliva:#656635;--ambar:#E89925;--off-white:#F2F1F0;\n  --bg:#656635;--on-bg:#F2F1F0;--on-bg-soft:#E4E1D3;--card:#FBF8EE;--field:#FFFFFF;--ink:#2B2A22;--soft:#5E5A4C;\n  --line:#E2DDCB;--brand:#656635;--brand-soft:#F0ECDD;--accent:#E89925;--on-accent:#2B2A22;\n  --ok:#38584B;--ok-ink:#F2F1F0;--mid:#E89925;--mid-ink:#2B2A22;--bad:#C3622D;--bad-ink:#FFFFFF;--neu:#88764F;--neu-ink:#FFFFFF;\n  --bad-text:#83462D;--alert-bg:#F7E6DA;--alert-line:#E9C9B4;--alert-ink:#83462D;\n  --f-display:\"Optima LT Pro\",Optima,\"Optima nova\",\"Marcellus\",Candara,serif;\n  --f-body:Garet,\"Figtree\",system-ui,-apple-system,\"Segoe UI\",sans-serif;\n  --shadow-card:0 2px 10px rgba(43,42,34,.10);color-scheme:light;\n}\n*{box-sizing:border-box}\nbody{margin:0;background:var(--bg);color:var(--ink);font-family:var(--f-body);-webkit-font-smoothing:antialiased;font-size:15px}\n.marca-faixa{background:var(--oliva);color:var(--off-white);padding:18px 16px 4px;text-align:center}\n.marca-faixa svg{width:18px;height:26px;display:block;margin:0 auto 8px;fill:var(--ambar)}\n.marca-faixa .nome{font:500 18px/1.15 var(--f-display);margin:0}\n.marca-faixa .cargo{font:600 10px/1.4 var(--f-body);letter-spacing:.14em;text-transform:uppercase;color:var(--on-bg-soft);margin-top:3px}\n.wrap{max-width:640px;margin:0 auto;padding:12px 16px 48px}\n.wrap.largo{max-width:1120px}\nh1{font:500 26px/1.15 var(--f-display);color:var(--on-bg);margin:10px 0 2px}\n.sub-h{color:var(--on-bg-soft);margin:0 0 12px;line-height:1.4}\nh2{font:500 19px/1.2 var(--f-display);margin:0 0 10px;color:var(--ink)}\n.card{background:var(--card);border-radius:16px;padding:14px 16px;margin:12px 0;box-shadow:var(--shadow-card)}\n.muted{color:var(--soft);font-size:13px}\n.row{display:flex;align-items:center;justify-content:space-between;gap:10px}\n.nav{display:flex;align-items:center;gap:8px;margin:6px 0 4px}\n.nav button,.btn2{border:0;border-radius:999px;background:var(--brand-soft);color:var(--ink);font:700 15px var(--f-body);padding:9px 14px;cursor:pointer}\n.nav input{flex:1;min-width:0;border:0;border-radius:999px;padding:9px 14px;font:700 15px var(--f-body);background:var(--card);color:var(--ink);text-align:center}\n.tabs{display:flex;gap:6px;margin:10px 0 2px}\n.tabs button{flex:1;border:2px solid transparent;border-radius:999px;padding:9px;font:700 14px var(--f-body);background:rgba(255,255,255,.12);color:var(--on-bg);cursor:pointer}\n.tabs button.on{background:var(--card);color:var(--ink)}\n.btn{border:0;border-radius:999px;padding:12px 18px;font:800 15px var(--f-body);background:var(--accent);color:var(--on-accent);cursor:pointer}\n.btn:disabled{opacity:.55}\n.btn.full{width:100%}\n.lnk{background:none;border:0;color:var(--bad-text);font:700 14px var(--f-body);cursor:pointer;padding:6px 0;text-decoration:underline}\n.chip{display:inline-block;border-radius:999px;padding:3px 10px;font-size:12px;font-weight:700;white-space:nowrap}\n.c-ok{background:var(--ok);color:var(--ok-ink)}.c-mid{background:var(--mid);color:var(--mid-ink)}\n.c-bad{background:var(--bad);color:var(--bad-ink)}.c-neu{background:var(--brand-soft);color:var(--soft)}\n.opts{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0 10px}\n.opts button{border:2px solid var(--line);background:var(--field);color:var(--ink);border-radius:12px;padding:8px 11px;font:700 14px var(--f-body);cursor:pointer}\n.opts button.on{border-color:var(--accent);background:var(--accent);color:var(--on-accent)}\nlabel.l{display:block;font-weight:700;font-size:13px;color:var(--soft);margin:8px 0 4px}\ninput.f,select.f,textarea.f{width:100%;border:1px solid var(--line);border-radius:10px;padding:10px;font:15px var(--f-body);background:var(--field);color:var(--ink)}\ntextarea.f{min-height:54px}\n.pg{display:grid;grid-template-columns:96px 1fr 90px;gap:6px;align-items:center;margin:4px 0}\n.pg .nm{font-weight:700}\n.alerta{background:var(--alert-bg);border:1px solid var(--alert-line);color:var(--alert-ink);border-radius:12px;padding:10px 12px;margin:10px 0;font-size:14px;line-height:1.45}\ntable{width:100%;border-collapse:collapse;font-size:14px}\nth,td{padding:7px 6px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}\nth{font-size:12px;color:var(--soft);text-transform:uppercase;letter-spacing:.04em}\ntd.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}\ntr.tot td{font-weight:800;border-bottom:0}\n.carregando{color:var(--on-bg-soft);text-align:center;padding:30px}\n.erro{background:var(--card);color:var(--bad-text);border-radius:12px;padding:14px;margin:12px 0}\n.rol{overflow-x:auto}\n.ba{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:6px 12px;padding:9px 0;border-bottom:1px solid var(--line)}\n.ba-nome{flex:1 1 260px;min-width:0}\n.ba-acao{display:flex;gap:6px;flex-wrap:wrap}\nfooter{margin-top:22px;text-align:center;font-size:12px;color:var(--on-bg-soft)}\n</style>\n<script>\nconst call = (fn, ...a) => new Promise((ok, no) => google.script.run.withSuccessHandler(ok).withFailureHandler(no)[fn](...a));\nconst R$ = (v) => (v || 0).toLocaleString(\"pt-BR\", { style: \"currency\", currency: \"BRL\" });\nconst esc = (s) => String(s == null ? \"\" : s).replace(/[&<>\"']/g, (c) => ({ \"&\": \"&amp;\", \"<\": \"&lt;\", \">\": \"&gt;\", '\"': \"&quot;\", \"'\": \"&#39;\" }[c]));\nconst dBR = (d) => d ? d.slice(8, 10) + \"/\" + d.slice(5, 7) + \"/\" + d.slice(0, 4) : \"\";\nconst SEMANA = [\"domingo\", \"segunda\", \"terça\", \"quarta\", \"quinta\", \"sexta\", \"sábado\"];\nconst diaSemana = (d) => SEMANA[new Date(d + \"T12:00:00\").getDay()];\nconst somaDias = (d, n) => { const x = new Date(d + \"T12:00:00\"); x.setDate(x.getDate() + n); return x.toISOString().slice(0, 10); };\nconst somaMes = (m, n) => { const x = new Date(m + \"-15T12:00:00\"); x.setMonth(x.getMonth() + n); return x.toISOString().slice(0, 7); };\nconst MESES = [\"janeiro\", \"fevereiro\", \"março\", \"abril\", \"maio\", \"junho\", \"julho\", \"agosto\", \"setembro\", \"outubro\", \"novembro\", \"dezembro\"];\nconst mesBR = (m) => MESES[+m.slice(5, 7) - 1] + \" de \" + m.slice(0, 4);\nconst num = (s) => { s = String(s || \"\").replace(/[^\\d,.]/g, \"\"); if (s.includes(\",\")) s = s.replace(/\\./g, \"\").replace(\",\", \".\"); else if (/^\\d{1,3}(\\.\\d{3})+$/.test(s)) s = s.replace(/\\./g, \"\"); const n = parseFloat(s); return isFinite(n) ? n : 0; };\nconst fmtIn = (v) => v ? String(v.toFixed(2)).replace(\".\", \",\") : \"\";\nconst waLink = (tel, nome) => {\n  let t = String(tel || \"\").replace(/\\D/g, \"\"); if (!t) return \"\";\n  if (!t.startsWith(\"55\")) t = \"55\" + t;\n  const msg = \"Olá, \" + String(nome || \"\").split(\" \")[0] + \"! Aqui é do consultório do Dr. Luis Felipe Jardim. Sentimos sua falta na última consulta. Está tudo bem com você? Quer que a gente remarque um horário?\";\n  return \"https://wa.me/\" + t + \"?text=\" + encodeURIComponent(msg);\n};\n/** Lista de busca ativa (pacientes com 2+ faltas em 180 dias). */\nfunction htmlBuscaAtiva(lista, comClinica) {\n  if (!lista.length) return '<p class=\"muted\">Nenhum paciente com 2 ou mais faltas nos últimos 6 meses.</p>';\n  return lista.map((p) => {\n    const sit = p.proxima ? '<span class=\"chip c-mid\">próxima ' + dBR(p.proxima.slice(0, 10)).slice(0, 5) + \" \" + esc(p.proxima.slice(11)) + \"</span>\"\n      : '<span class=\"chip c-bad\">sem retorno marcado</span>';\n    const wa = waLink(p.tel, p.paciente);\n    return '<div class=\"ba\"><div class=\"ba-nome\"><b>' + esc(p.paciente) + \"</b>\" + (comClinica ? ' <span class=\"muted\">· ' + esc(p.clinica) + \"</span>\" : \"\") +\n      '<div class=\"muted\">faltou ' + p.faltas + \" de \" + (p.faltas + p.atendidos) + \" · última falta \" + dBR(p.ultimaFalta).slice(0, 5) + (p.faltouDuasUltimas ? \" · faltou nas 2 últimas\" : \"\") + \"</div></div>\" +\n      '<div class=\"ba-acao\">' + sit + (wa ? ' <a class=\"chip c-ok\" style=\"text-decoration:none\" href=\"' + wa + '\" target=\"_blank\" rel=\"noopener\">WhatsApp</a>' : \"\") + \"</div></div>\";\n  }).join(\"\");\n}\nconst msgErro = (e) => (e && e.message ? e.message : String(e)).replace(/^Exception: /, \"\").replace(/^Error: /, \"\");\n</script>\n"
};
