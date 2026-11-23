/**
 * Pré-consulta → prontuário do V Saúde
 * ------------------------------------
 * Roda dentro da planilha "Pré-consulta por voz · recebimento (respostas)".
 * A cada resposta enviada pela página de pré-consulta:
 *   1. lê o CPF do cabeçalho;
 *   2. procura o paciente no V Saúde pelo CPF;
 *   3. gera um PDF organizado e anexa na pasta de Arquivos do paciente;
 *   4. escreve o resultado na coluna "Status V Saúde" da planilha.
 * Se o paciente não for encontrado (ou houver mais de um), envia um e-mail
 * de aviso e deixa a linha marcada como pendente; a cada hora tenta de novo.
 *
 * Instalação: veja integracao/COMO-INSTALAR.md
 */

const VS_BASE = "https://public-api.vsaude.com.br/api/services/app";
const PASTA_RAIZ = "00000000-0000-0000-0000-000000000000";
const COL_STATUS = "Status V Saúde";
const FREQ = ["Nada", "Muito leve", "Leve", "Moderado", "Grave"];

// ===================== GATILHOS =====================

/** Rodar UMA vez, pelo menu Executar, depois de guardar a chave da API. */
function instalar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("aoReceberResposta").forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger("reprocessarPendentes").timeBased().everyHours(1).create();
  colunaStatus_();
  if (!chave_()) throw new Error("Falta guardar a chave: Configurações do projeto → Propriedades do script → VSAUDE_API_KEY");
  vs_("POST", "/HealthCareUnitService/GetAll", {}); // testa a chave
  Logger.log("Instalado. A chave do V Saúde funciona.");
}

function aoReceberResposta(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { processarLinha_(e.range.getRow()); } finally { lock.releaseLock(); }
}

/** Tenta de novo as linhas ainda não anexadas (roda a cada hora). */
function reprocessarPendentes() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return;
  try {
    const sh = aba_(), col = colunaStatus_(), n = sh.getLastRow();
    for (let r = 2; r <= n; r++) {
      const st = String(sh.getRange(r, col).getValue());
      if (!st.startsWith("✅") && !st.startsWith("⛔")) processarLinha_(r, true);
    }
  } finally { lock.releaseLock(); }
}

// ===================== PROCESSAMENTO =====================

function processarLinha_(linha, silencioso) {
  const sh = aba_(), col = colunaStatus_();
  const [quando, cab, corpo] = sh.getRange(linha, 1, 1, 3).getValues()[0];
  if (!cab || !corpo) return;
  const status = msg => sh.getRange(linha, col).setValue(msg);
  try {
    const d = interpretar(String(cab), String(corpo));
    if (!d.cpf) { status("⚠ Sem CPF na resposta"); return; }
    const achados = buscarPorCpf_(d.cpf);
    if (achados.length !== 1) {
      const msg = achados.length ? `⚠ ${achados.length} cadastros com este CPF` : "⚠ CPF não encontrado no V Saúde";
      const anterior = String(sh.getRange(linha, col).getValue());
      status(msg + " · tentando de hora em hora");
      if (!silencioso || !anterior.startsWith("⚠")) avisar_(d, msg, linha);
      return;
    }
    const p = achados[0];
    d.vsaude = { id: p.id, numero: p.sequentialNumber, nome: p.name };
    const data = Utilities.formatDate(quando instanceof Date ? quando : new Date(), "America/Sao_Paulo", "dd-MM-yyyy");
    const nomeArq = `Pré-consulta – ${data}.pdf`;
    const pdf = HtmlService.createHtmlOutput(montarHtml(d)).getAs("application/pdf").setName(nomeArq);
    enviarArquivo_(p.id, nomeArq, pdf);
    status(`✅ Anexado em ${Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM HH:mm")} · paciente nº ${p.sequentialNumber}`);
  } catch (err) {
    status("⚠ Erro: " + String(err.message || err).slice(0, 180) + " · tentando de hora em hora");
    if (!silencioso) avisar_(null, String(err.message || err), linha);
  }
}

// ===================== V SAÚDE =====================

function chave_() { return PropertiesService.getScriptProperties().getProperty("VSAUDE_API_KEY"); }

function vs_(metodo, caminho, corpo, extra) {
  const op = Object.assign({ method: metodo, muteHttpExceptions: true, headers: { "VSAUDE-API-KEY": chave_() } }, extra || {});
  if (corpo !== undefined && !op.payload) { op.contentType = "application/json"; op.payload = JSON.stringify(corpo); }
  const r = UrlFetchApp.fetch(VS_BASE + caminho, op);
  let j = null; try { j = JSON.parse(r.getContentText()); } catch (e) {}
  if (r.getResponseCode() >= 300 || !j || j.success === false)
    throw new Error(`V Saúde ${r.getResponseCode()}: ${(j && j.error && j.error.message) || r.getContentText().slice(0, 150)}`);
  return j.result;
}

function buscarPorCpf_(cpf) {
  const fmt = cpf.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, "$1.$2.$3-$4");
  const vistos = {}, out = [];
  [fmt, cpf].forEach(k => {
    (vs_("POST", "/PatientService/Search?keyword=" + encodeURIComponent(k)) || []).forEach(p => {
      if (String(p.personalIdentifier || "").replace(/\D/g, "") === cpf && !vistos[p.id]) { vistos[p.id] = 1; out.push(p); }
    });
  });
  return out;
}

function enviarArquivo_(pacienteId, nome, blob) {
  const q = `?parent=${PASTA_RAIZ}&ownerPatient=${pacienteId}&fileName=${encodeURIComponent(nome)}`;
  return vs_("POST", "/FilesService/Upload" + q, undefined, { payload: { file: blob } });
}

// ===================== PLANILHA E AVISO =====================

function aba_() { return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }

function colunaStatus_() {
  const sh = aba_(), ult = Math.max(sh.getLastColumn(), 3);
  const cab = sh.getRange(1, 1, 1, ult).getValues()[0];
  const i = cab.indexOf(COL_STATUS);
  if (i >= 0) return i + 1;
  sh.getRange(1, ult + 1).setValue(COL_STATUS).setFontWeight("bold");
  return ult + 1;
}

function avisar_(d, motivo, linha) {
  const para = Session.getEffectiveUser().getEmail(); if (!para) return;
  const quem = d ? `${d.nome || "(sem nome)"} · CPF ${d.cpfFmt || "—"}` : "(não identificado)";
  MailApp.sendEmail(para, `Pré-consulta não anexada ao V Saúde: ${d ? d.nome : ""}`,
    `${motivo}\n\nPaciente: ${quem}\nLinha da planilha: ${linha}\n\n` +
    `O que fazer: confira se o paciente está cadastrado no V Saúde com este CPF (ou corrija o CPF no cadastro). ` +
    `A planilha tenta de novo a cada hora e anexa sozinha quando encontrar.\n\n` +
    `Planilha: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
}

// ===================== LEITURA DA RESPOSTA =====================
// O cabeçalho e o corpo vêm da página preconsulta.html (função enviar).

function interpretar(cab, corpo) {
  const d = { roteiro: (cab.match(/^\[([^\]]+)\]/) || [])[1] || "", secoes: [], risco: /⚠ ATENÇÃO/.test(corpo) };
  const cpf = (cab.match(/CPF (\d{11})/) || [])[1] || "";
  d.cpf = cpf; d.cpfFmt = cpf.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, "$1.$2.$3-$4");
  d.tempo = (corpo.match(/Tempo de preenchimento: ([^\n]+)/) || [])[1] || "";
  d.enviado = (corpo.match(/\n(\d{2}\/\d{2}\/\d{4},? \d{2}:\d{2})/) || [])[1] || "";
  corpo.split(/\n## /).slice(1).forEach(b => {
    const k = b.indexOf("\n"), titulo = (k < 0 ? b : b.slice(0, k)).trim(), texto = (k < 0 ? "" : b.slice(k + 1)).trim();
    d.secoes.push({ titulo, texto: texto === "—" ? "" : texto });
  });
  const campo = (sec, rot) => { const s = d.secoes.find(x => x.titulo === sec); if (!s) return "";
    const l = s.texto.split("\n").find(x => x.startsWith(rot + ":")); return l ? l.slice(rot.length + 1).trim() : ""; };
  d.ident = {}; const id = d.secoes.find(s => s.titulo === "Identificação");
  if (id) id.texto.split("\n").forEach(l => { const k = l.indexOf(":"); if (k > 0) d.ident[l.slice(0, k).trim()] = l.slice(k + 1).trim(); });
  d.nome = Object.keys(d.ident).filter(k => /^Nome completo/.test(k)).map(k => d.ident[k])[0] || "";
  d.nasc = Object.keys(d.ident).filter(k => /^Data de nascimento/.test(k)).map(k => d.ident[k])[0] || "";
  d.tel = Object.keys(d.ident).filter(k => /^Celular/.test(k)).map(k => d.ident[k])[0] || "";
  d.respondente = d.ident["Quem está respondendo?"] || (d.ident["Seu nome (responsável)"] ? `${d.ident["Seu nome (responsável)"]} (${d.ident["Parentesco (mãe, pai, avó, outro)"] || "responsável"})` : "O próprio paciente");
  d.cuidador = d.ident["Nome e parentesco do familiar ou cuidador (se houver)"] || "";
  const imp = d.secoes.find(s => /^Quanto atrapalha|^Áreas:/m.test(s.texto));
  d.impacto = imp ? { titulo: imp.titulo, areas: (imp.texto.match(/Áreas: ([^\n]*)/) || [])[1] || "", nota: (imp.texto.match(/\(0–10\): ([^\n]*)/) || [])[1] || "" } : null;
  d.motivacao = { imp: campo("Motivação", "Importância de resolver agora (0–10)"), conf: campo("Motivação", "Confiança de mudar com ajuda (0–10)") };
  const ras = d.secoes.find(s => s.texto && s.texto.split("\n").every(l => /: (Nada|Muito leve|Leve|Moderado|Grave|—)$/.test(l)));
  d.rastreio = ras ? { titulo: ras.titulo, itens: ras.texto.split("\n").map(l => { const k = l.lastIndexOf(": "); return { q: l.slice(0, k), v: FREQ.indexOf(l.slice(k + 2)) }; }) } : null;
  if (d.rastreio) d.rastreio.itens.forEach(it => it.risco = /ferir a si|em morrer|se machuca de prop/i.test(it.q));
  d.riscoItens = d.rastreio ? d.rastreio.itens.filter(it => it.risco && it.v > 0) : [];
  const idade = d.nasc.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (idade) { const n = new Date(+idade[3], +idade[2] - 1, +idade[1]), h = new Date(); let a = h.getFullYear() - n.getFullYear();
    if (h < new Date(h.getFullYear(), n.getMonth(), n.getDate())) a--; d.idade = a; }
  return d;
}

// ===================== PDF =====================

function esc_(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

function montarHtml(d) {
  const e = esc_, curto = q => q.split(/,| \(/)[0];
  const abertas = d.secoes.filter(s => s.titulo !== "Identificação" && s !== (d.rastreio && d.secoes.find(x => x.titulo === d.rastreio.titulo)) &&
    !(d.impacto && s.titulo === d.impacto.titulo) && s.titulo !== "Motivação");
  const destaques = d.rastreio ? d.rastreio.itens.filter(it => it.v >= 3) : [];
  const frases = abertas.map(s => s.texto).join(" ").split(/(?<=[.!?])\s+/)
    .filter(f => /suicid|me matar|tirar a (minha |própria )?vida|sumir|morrer|me ferir|me cortar|acabar com tudo/i.test(f)).slice(0, 2);
  return `<!doctype html><html><head><meta charset="utf-8"><style>
  @page{size:A4;margin:14mm 14mm 16mm}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;line-height:1.45;color:#2B2A22;margin:0}
  .cab{border-bottom:2.5px solid #656635;padding-bottom:5px;overflow:hidden}
  .cab h1{font-family:Georgia,serif;font-weight:normal;font-size:20pt;color:#656635;margin:0;float:left}
  .tag{background:#656635;color:#fff;font-size:8pt;font-weight:bold;padding:2px 8px;border-radius:9px;margin-left:8px;vertical-align:4px}
  .dr{float:right;text-align:right;font-size:8pt;color:#5f5d4c;line-height:1.35}
  .id{background:#FBF8EE;border:1px solid #e2dccb;border-radius:7px;padding:8px 11px;margin:9px 0}
  .id table{width:100%;border-collapse:collapse} .id td{padding:2px 6px 2px 0;font-size:9.3pt;vertical-align:top;border:0}
  .id b{display:block;font-size:7pt;color:#5f5d4c;text-transform:uppercase;font-weight:bold}
  .nome{font-size:12.5pt;font-weight:bold}
  .alerta{border:1.5px solid #C3622D;background:#fbeee6;color:#83462D;border-radius:7px;padding:7px 11px;margin:7px 0;font-size:9.3pt}
  h2{font-family:Georgia,serif;font-weight:normal;font-size:13pt;color:#656635;margin:13px 0 5px;border-bottom:1px solid #e2dccb;padding-bottom:2px}
  .vr td{border:1px solid #e2dccb;border-radius:6px;padding:6px 9px;vertical-align:top;font-size:9.3pt;width:33%}
  .vr b{display:block;font-size:7pt;color:#5f5d4c;text-transform:uppercase}
  .big{font-size:16pt;font-weight:bold;color:#656635}
  table.r{width:100%;border-collapse:collapse;font-size:8.8pt}
  table.r th{font-size:7pt;color:#5f5d4c;text-transform:uppercase;padding:3px 2px;border-bottom:1.5px solid #656635}
  table.r td{padding:3px 4px;border-bottom:1px solid #eee8d8} td.c{width:50px;text-align:center;color:#656635}
  tr.g3 td{background:#f6efdc} tr.g4 td{background:#f1e2c0;font-weight:bold} tr.rk td{background:#fbeee6;color:#83462D;font-weight:bold}
  .leg{font-size:7.8pt;color:#5f5d4c;margin-top:3px}
  h3{font-size:10.3pt;color:#38584B;margin:9px 0 2px} p{margin:0 0 3px;text-align:justify}
  .vazio{color:#5f5d4c;font-style:italic}
  .nota{font-size:8pt;color:#5f5d4c;border-top:1px solid #e2dccb;margin-top:12px;padding-top:5px}
  </style></head><body>
  <div class="cab"><h1>Pré-consulta <span class="tag">${e((d.roteiro || "").toUpperCase())}</span></h1>
  <div class="dr">Dr. Luis Felipe Jardim · Psiquiatra<br>CRM-SC 12767 · RQE 22739<br>Recebida em ${e(d.enviado)}</div></div>
  <div class="id"><div class="nome">${e(d.nome)}</div><table><tr>
   <td><b>CPF</b>${e(d.cpfFmt)}</td><td><b>Nascimento · idade</b>${e(d.nasc)}${d.idade != null ? " · " + d.idade + " anos" : ""}</td><td><b>Celular</b>${e(d.tel)}</td></tr><tr>
   <td><b>Respondido por</b>${e(d.respondente)}${d.cuidador ? "<br>" + e(d.cuidador) : ""}</td>
   <td><b>Cadastro V Saúde</b>${d.vsaude ? "Paciente nº " + e(d.vsaude.numero) + " · localizado pelo CPF" : "—"}</td>
   <td><b>Tempo de preenchimento</b>${e(d.tempo || "—")}</td></tr></table></div>
  ${d.risco || d.riscoItens.length ? `<div class="alerta"><b>⚠ Sinal de risco.</b> ${d.riscoItens.map(it => `Rastreio: “${e(it.q)}” = <b>${FREQ[it.v]}</b>.`).join(" ")}
    ${frases.length ? " No relato: " + frases.map(f => "“" + e(f) + "”").join(" ") : ""} A orientação do CVV 188 / SAMU 192 foi exibida na tela.</div>` : ""}
  <h2>Visão rápida</h2>
  <table class="vr" style="width:100%;border-collapse:separate;border-spacing:6px 0"><tr>
   <td><b>Quanto atrapalha hoje</b>${d.impacto ? `<span class="big">${e(d.impacto.nota || "—")}</span> / 10<br>${e(d.impacto.areas.split(", ").join(" · "))}` : "—"}</td>
   <td><b>Motivação</b>Importância de resolver agora: <strong>${e(d.motivacao.imp || "—")}</strong>/10<br>Confiança de mudar com ajuda: <strong>${e(d.motivacao.conf || "—")}</strong>/10</td>
   <td><b>Rastreio: moderado ou grave</b>${destaques.length ? destaques.map(it => "• " + e(curto(it.q))).join("<br>") : "Nenhum item"}</td></tr></table>
  ${d.rastreio ? `<h2>Rastreio · ${e(d.rastreio.titulo.toLowerCase())}</h2>
  <table class="r"><tr><th></th><th style="text-align:left">Item</th>${FREQ.map(f => `<th>${f}</th>`).join("")}</tr>
  ${d.rastreio.itens.map((it, i) => `<tr class="${it.risco && it.v > 0 ? "rk" : it.v === 4 ? "g4" : it.v === 3 ? "g3" : ""}"><td style="color:#5f5d4c;text-align:right">${i + 1}</td><td>${e(it.q)}</td>${FREQ.map((_, j) => `<td class="c">${it.v === j ? "●" : ""}</td>`).join("")}</tr>`).join("")}</table>
  <div class="leg">Muito leve = 1–2 dias · Leve = vários dias · Moderado = mais da metade dos dias · Grave = quase todos os dias. Itens sem marcação: não respondidos.</div>` : ""}
  <h2>Respostas abertas</h2>
  <div class="leg" style="margin:0 0 4px">Texto ditado ou digitado por quem respondeu e revisado antes do envio. A oralidade foi preservada.</div>
  ${abertas.map(s => `<h3>${e(s.titulo)}</h3>${s.texto ? s.texto.split("\n").map(l => `<p>${e(l)}</p>`).join("") : '<p class="vazio">Não respondido.</p>'}`).join("")}
  <div class="nota">Documento gerado automaticamente a partir da pré-consulta on-line e anexado ao prontuário. Conteúdo sigiloso, protegido pelo sigilo médico e pela LGPD (Lei 13.709/2018).</div>
  </body></html>`;
}

if (typeof module !== "undefined") module.exports = { interpretar, montarHtml }; // testes locais
