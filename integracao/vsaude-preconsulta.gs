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
    // Acompanhamento da depressão: o check-in só fica registrado na planilha (entra no próximo PDF do PHQ-9);
    // sinal de risco (PHQ-9 item 9, ou "pensamentos de morte" no check-in) gera e-mail imediato.
    if (d.roteiro === "CHECK-IN") {
      if (d.risco && !silencioso) { try { localizar_(d); } catch (e) {} alertarRisco_(d, linha); } // localizar completa nome, CPF e celular
      status("✅ Check-in registrado · entra no próximo PDF do PHQ-9"); return;
    }
    const loc = localizar_(d);
    if (d.roteiro === "PHQ-9" && d.risco && !silencioso) alertarRisco_(d, linha);
    if (!loc.p) {
      const msg = "⚠ " + loc.erro;
      const anterior = String(sh.getRange(linha, col).getValue());
      status(msg + " · tentando de hora em hora");
      if (!silencioso || !anterior.startsWith("⚠")) avisar_(d, msg, linha);
      return;
    }
    const p = loc.p;
    if (d.roteiro === "PHQ-9") d.historico = historicoHumor_(sh, d.cpf, d.nome, linha, d.codigo);
    d.vsaude = { id: p.id, numero: p.sequentialNumber, nome: p.name, via: loc.via };
    const data = Utilities.formatDate(ehData_(quando) ? quando : new Date(), "America/Sao_Paulo", "dd-MM-yyyy");
    const nomeArq = `${{ "DIVA-5": "DIVA-5 (TDAH)", "PHQ-9": "PHQ-9 (depressão)" }[d.roteiro] || "Pré-consulta"} – ${data}.pdf`;
    const pdf = HtmlService.createHtmlOutput(montarHtml(d)).getAs("application/pdf").setName(nomeArq);
    enviarArquivo_(p.id, nomeArq, pdf);
    status(`✅ Anexado em ${Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM HH:mm")} · ${p.sequentialNumber ? "paciente nº " + p.sequentialNumber + " · " : ""}localizado pelo ${loc.via}${loc.obs ? " · " + loc.obs : ""}`);
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

/**
 * Localiza o paciente no V Saúde.
 * 1) pelo CPF; 2) se o CPF não estiver no cadastro, pelo NOME, conferido com a DATA DE NASCIMENTO.
 * Só anexa quando sobra exatamente um cadastro; homônimos sem como desempatar ficam pendentes.
 */
function localizar_(d) {
  // Link pessoal (gerado pelo consultório para um cadastro do V Saúde): vale mais que qualquer busca.
  if (d.codigo) {
    const l = linkPessoal_(d.codigo);
    if (l && l.cpf) {
      const a = buscarPorCpf_(l.cpf);
      if (a.length !== 1) return { erro: a.length ? `${a.length} cadastros com o CPF do link ${d.codigo}` : `CPF do link ${d.codigo} não encontrado no V Saúde` };
      l.id = a[0].id;
    }
    if (l) {
      const p = vs_("GET", "/PatientService/Get?Id=" + l.id);
      if (!p) return { erro: `Link ${d.codigo}: cadastro não existe mais no V Saúde` };
      const cpfCad = String(p.personalIdentifier || "").replace(/\D/g, ""), b = String(p.birthday || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
      // O PDF usa os dados do cadastro.
      d.nome = p.name || d.nome; if (cpfCad) { d.cpf = cpfCad; d.cpfFmt = cpfCad.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, "$1.$2.$3-$4"); }
      if (b) { d.nasc = `${b[3]}/${b[2]}/${b[1]}`; const h = new Date(), n = new Date(+b[1], +b[2] - 1, +b[3]); d.idade = h.getFullYear() - n.getFullYear() - (h < new Date(h.getFullYear(), n.getMonth(), n.getDate()) ? 1 : 0); } if (!d.tel && p.phone) d.tel = String(p.phone).replace(/^\+55/, "");
      return { p: { id: p.id || l.id, name: p.name, sequentialNumber: p.sequentialNumber, personalIdentifier: p.personalIdentifier }, via: "link pessoal" };
    }
  }
  if (d.cpf) {
    const a = buscarPorCpf_(d.cpf);
    if (a.length === 1) return { p: a[0], via: "CPF" };
    if (a.length > 1) return { erro: `${a.length} cadastros com este CPF` };
  }
  const nome = norm_(d.nome);
  if (!nome) return { erro: d.cpf ? "CPF não encontrado no V Saúde e resposta sem nome" : "Resposta sem CPF e sem nome" };
  const cand = {};
  const buscar = k => (vs_("POST", "/PatientService/Search?keyword=" + encodeURIComponent(k)) || []).forEach(p => (cand[p.id] = p));
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(d.nasc || "")) buscar(d.nasc);
  buscar(d.nome.trim());
  let iguais = Object.keys(cand).map(k => cand[k]).filter(p => mesmoNome_(norm_(p.name), nome));
  if (!iguais.length) return { erro: "Paciente não encontrado no V Saúde (nem pelo CPF, nem pelo nome)" };
  iguais.forEach(p => (p.nasc_ = nascDe_(p.id)));
  if (d.nasc) {
    const conf = iguais.filter(p => p.nasc_ === d.nasc);
    if (conf.length === 1) return achado_(conf[0], "nome e data de nascimento", d);
    if (conf.length > 1) return { erro: `${conf.length} cadastros com o mesmo nome e nascimento` };
    iguais = iguais.filter(p => !p.nasc_); // cadastro sem data: só vale se o nome for idêntico e único
  }
  const exatos = iguais.filter(p => norm_(p.name) === nome);
  if (exatos.length === 1) return achado_(exatos[0], "nome (cadastro sem data de nascimento)", d);
  return { erro: iguais.length > 1 ? `${iguais.length} cadastros com nome parecido; confira o CPF no cadastro` : "Nome encontrado, mas a data de nascimento não confere" };
}
function achado_(p, via, d) {
  const cpfCad = String(p.personalIdentifier || "").replace(/\D/g, "");
  return { p, via, obs: cpfCad && d.cpf && cpfCad !== d.cpf ? "CPF do cadastro difere do informado" : !cpfCad && d.cpf ? "cadastro sem CPF: vale completar" : "" };
}
const PARTICULAS_ = { de: 1, da: 1, do: 1, das: 1, dos: 1, e: 1 };
function norm_(s) {
  return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z ]/g, " ")
    .split(/\s+/).filter(w => w && !PARTICULAS_[w]).join(" ");
}
/** Mesmo nome: idêntico, ou mesmo primeiro e último nome e um contido no outro (nome do meio a mais ou a menos). */
function mesmoNome_(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  const x = a.split(" "), y = b.split(" ");
  if (x[0] !== y[0] || x[x.length - 1] !== y[y.length - 1]) return false;
  const [cur, lon] = x.length <= y.length ? [x, y] : [y, x];
  return cur.every(w => lon.indexOf(w) >= 0);
}
function nascDe_(id) {
  try { const b = String((vs_("GET", "/PatientService/Get?Id=" + id) || {}).birthday || ""); const m = b.match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : ""; }
  catch (e) { return ""; }
}

/** Aba "Links" da planilha: código | id do paciente no V Saúde | nome | app | criado em. */
function linkPessoal_(codigo) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Links");
  if (!aba || aba.getLastRow() < 2) return null;
  const l = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues().find(r => String(r[0]).trim().toUpperCase() === codigo);
  if (!l || !l[1]) return null;
  const v = String(l[1]).trim(), dig = v.replace(/\D/g, "");
  // Coluna B: o ID do cadastro (busca no app) ou, quando o link foi criado só com o CPF, o CPF.
  return /^[0-9a-f-]{36}$/i.test(v) ? { id: v } : dig.length === 11 ? { cpf: dig } : null;
}

const ehData_ = v => Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v);

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
  MailApp.sendEmail(para, `${d && (d.roteiro === "DIVA-5" || d.roteiro === "PHQ-9") ? d.roteiro : "Pré-consulta"} não anexada ao V Saúde: ${d ? d.nome : ""}`,
    `${motivo}\n\nPaciente: ${quem}\nLinha da planilha: ${linha}\n\n` +
    `O que fazer: confira se o paciente está cadastrado no V Saúde com este CPF, ou com o mesmo nome e data de nascimento (corrija ou complete o cadastro). ` +
    `A planilha tenta de novo a cada hora e anexa sozinha quando encontrar.\n\n` +
    `Planilha: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
}

// ===================== LEITURA DA RESPOSTA =====================
// O cabeçalho e o corpo vêm da página preconsulta.html (função enviar).

function interpretar(cab, corpo) {
  const d = { roteiro: (cab.match(/^\[([^\]]+)\]/) || [])[1] || "", secoes: [], risco: /⚠ ATENÇÃO/.test(corpo) };
  const cpf = (cab.match(/CPF (\d{11})/) || [])[1] || "";
  d.cpf = cpf; d.cpfFmt = cpf.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, "$1.$2.$3-$4");
  d.codigo = (cab.match(/código ([A-Z0-9]{6,16})/) || [])[1] || "";
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

// ===================== DIVA-5: GRÁFICO DO RESULTADO =====================
// Tabelas simples (o conversor de PDF do Google não roda JavaScript).
// Cor = fase: azul = hoje (adulto), terracota = infância (5–12 anos). Paleta conferida para daltonismo.
const DIVA_COR = { ad: "#2F55B5", inf: "#C3622D" };
function divaDados_(d) {
  const it = d.secoes.filter(s => /^(A|H\/I )\d · /.test(s.titulo)).map(s => ({
    cod: s.titulo.split(" · ")[0], nome: s.titulo.split(" · ")[1],
    ad: /^Adulto: SIM/m.test(s.texto), inf: /^Infância: SIM/m.test(s.texto) }));
  const A = it.filter(x => x.cod[0] === "A"), H = it.filter(x => x.cod[0] === "H");
  const res = (d.secoes.find(s => /^Resumo para o médico/.test(s.titulo)) || {}).texto || "";
  const area = l => { const m = res.match(new RegExp("áreas com prejuízo: " + l + " (\\d)/5 \\(([^)]*)\\)")); return m ? { n: +m[1], quais: m[2] } : { n: 0, quais: "" }; };
  return { A, H, B: (res.match(/Critério B \(antes dos 12 anos\): ([^\n]*)/) || [])[1] || "—",
    Cad: area("adulto"), Cinf: area("infância") };
}
function divaHtml_(d) {
  const e = esc_, g = divaDados_(d);
  const n = (xs, f) => xs.filter(x => x[f]).length;
  const linhas = [["Déficit de atenção", g.A], ["Hiperatividade / impulsividade", g.H]];
  const fase = { ad: ["Hoje", 5], inf: ["Infância", 6] };
  const barra = (xs, f) => {
    const k = n(xs, f), lim = fase[f][1], ok = k >= lim;
    return `<tr><td class="dl">${fase[f][0]}</td><td><table class="seg"><tr>${[1,2,3,4,5,6,7,8,9].map(i =>
      `<td style="${i <= k ? "background:" + DIVA_COR[f] : "background:#ECEAE3"}${i === lim ? ";border-left:2.5px solid #2B2A22" : ""}">&nbsp;</td>`).join("")}</tr></table></td>
      <td class="dv"><b>${k}</b>/9</td><td class="dk">${ok ? "atinge o limiar (≥ " + lim + ")" : "abaixo do limiar (≥ " + lim + ")"}</td></tr>`;
  };
  const apres = (f, lim) => { const a = n(g.A, f) >= lim, h = n(g.H, f) >= lim;
    return a && h ? "combinada" : a ? "predominantemente desatenta" : h ? "predominantemente hiperativa/impulsiva" : "abaixo do limiar nas duas dimensões"; };
  const mapa = (xs) => `<table class="mapa"><tr><th></th>${xs.map(x => `<th title="${e(x.nome)}">${e(x.cod.replace("H/I ", "HI"))}</th>`).join("")}</tr>
    ${["ad", "inf"].map(f => `<tr><td class="dl">${fase[f][0]}</td>${xs.map(x => `<td class="c">${x[f] ? `<span style="color:${DIVA_COR[f]}">●</span>` : '<span style="color:#C9C5B8">○</span>'}</td>`).join("")}</tr>`).join("")}</table>
    <div class="leg">${xs.map(x => `${e(x.cod.replace("H/I ", "HI"))} ${e(x.nome.toLowerCase())}`).join(" · ")}</div>`;
  return `<h2>Resultado</h2>
  <div class="leg" style="margin:0 0 6px">Autorrelato do paciente; não é diagnóstico. Cada quadrado é um sintoma marcado como presente. A linha preta marca o limiar do DSM-5 (adulto ≥ 5; infância ≥ 6 sintomas).</div>
  <table class="dv-leg"><tr><td><span style="background:${DIVA_COR.ad}"></span> Hoje (últimos 6 meses)</td><td><span style="background:${DIVA_COR.inf}"></span> Infância (5 a 12 anos)</td><td><i></i> limiar DSM-5</td></tr></table>
  ${linhas.map(([t, xs]) => `<h3>${t}</h3><table class="graf">${barra(xs, "ad")}${barra(xs, "inf")}</table>`).join("")}
  <table class="sint"><tr>
   <td><b>Apresentação pelo autorrelato</b>Hoje: ${apres("ad", 5)}<br>Infância: ${apres("inf", 6)}</td>
   <td><b>Início antes dos 12 anos (B)</b>${e(g.B)}</td>
   <td><b>Prejuízo em 2 ou mais contextos (C/D)</b>Hoje: ${g.Cad.n}/5 áreas ${g.Cad.n >= 2 ? "· sim" : "· não"}${g.Cad.quais && g.Cad.quais !== "nenhuma" ? "<br><span class='mini'>" + e(g.Cad.quais) + "</span>" : ""}<br>Infância: ${g.Cinf.n}/5 áreas ${g.Cinf.n >= 2 ? "· sim" : "· não"}${g.Cinf.quais && g.Cinf.quais !== "nenhuma" ? "<br><span class='mini'>" + e(g.Cinf.quais) + "</span>" : ""}</td></tr></table>
  <h3>Mapa dos sintomas</h3>${mapa(g.A)}${mapa(g.H)}`;
}

// ===================== ACOMPANHAMENTO DA DEPRESSÃO (PHQ-9) =====================
const PHQ_CURTO = ["Pouco interesse ou prazer", "Para baixo, sem perspectiva", "Sono", "Cansaço, pouca energia", "Apetite",
  "Sentir-se mal consigo, fracasso", "Concentração", "Lentidão ou agitação", "Pensar em se ferir ou em morrer"];
const PHQ_FAIXAS = [[0, 4, "mínimos"], [5, 9, "leves"], [10, 14, "moderados"], [15, 19, "moderadamente graves"], [20, 27, "graves"]];
const PHQ_TONS = ["#DCE6E0", "#B4CBBE", "#86A796", "#56806D", "#2F5A48"]; // um tom só: claro = mínimo, escuro = grave
const phqFaixa_ = t => PHQ_FAIXAS.findIndex(([a, b]) => t >= a && t <= b);
/** Lê um registro (PHQ-9 ou check-in) a partir do texto enviado pela página. */
function registroHumor_(quando, cab, corpo) {
  const tipo = (String(cab).match(/^\[([^\]]+)\]/) || [])[1], c = String(corpo), g = re => (c.match(re) || [])[1];
  const r = { quando: ehData_(quando) ? quando : new Date(quando), tipo,
    humor: g(/Humor \(0–10\): (\d+)/), sono: g(/\nSono: ([^\n]*)/), remedio: g(/\nRemédios: ([^\n]*)/),
    efeitos: g(/\nEfeitos colaterais: ([^\n]*)/), morte: g(/Pensamentos de morte ou de se ferir: ([^\n]*)/),
    nota: (c.split("## Observações do paciente\n")[1] || "").trim(), risco: /⚠ ATENÇÃO/.test(c) };
  if (r.nota === "—") r.nota = "";
  if (tipo === "PHQ-9") {
    r.itens = [1, 2, 3, 4, 5, 6, 7, 8, 9].map(i => +((c.match(new RegExp("\\n" + i + "\\. [^\\n]*: (\\d) \\(")) || [])[1] || 0));
    r.total = +(g(/Total: (\d+)/) || 0); r.func = g(/## Dificuldade funcional\n([^\n]*)/);
  }
  return r;
}
/** Histórico do mesmo paciente na planilha (PHQ-9 e check-ins), até a linha atual. */
function historicoHumor_(sh, cpf, nome, ateLinha, codigo) {
  const n = Math.min(sh.getLastRow(), ateLinha);
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 3).getValues()
    .filter(([q, cab]) => /^\[(PHQ-9|CHECK-IN)\]/.test(String(cab)) && ((codigo && String(cab).indexOf("código " + codigo) >= 0) || (cpf && String(cab).indexOf("CPF " + cpf + " ") >= 0) || (!cpf && !codigo && norm_(String(cab).split(" · ")[0].replace(/^\[[^\]]+\] /, "")) === norm_(nome))))
    .map(([q, cab, corpo]) => registroHumor_(q, cab, corpo)).sort((a, b) => a.quando - b.quando);
}
function alertarRisco_(d, linha) {
  const para = Session.getEffectiveUser().getEmail(); if (!para) return;
  MailApp.sendEmail(para, `⚠ Sinal de risco · ${d.roteiro} · ${d.nome}`,
    `${d.nome} (CPF ${d.cpfFmt || "—"}, celular ${d.tel || "—"}) marcou pensamentos de morte ou de se ferir no ${d.roteiro === "PHQ-9" ? "PHQ-9 (item 9)" : "check-in"}.\n` +
    `A página mostrou ao paciente a orientação do CVV 188 / SAMU 192.\n\nLinha da planilha: ${linha}\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
}
function phqHtml_(d) {
  const e = esc_, data = q => Utilities.formatDate(q, "America/Sao_Paulo", "dd/MM/yy");
  const h = (d.historico || []).length ? d.historico : [], ps = h.filter(r => r.tipo === "PHQ-9");
  const atual = ps[ps.length - 1], base = ps[0], ant = ps[ps.length - 2];
  if (!atual) return "";
  const desde = ant ? h.filter(r => r.tipo === "CHECK-IN" && r.quando > ant.quando) : h.filter(r => r.tipo === "CHECK-IN");
  const varPct = base && base.total ? Math.round((atual.total - base.total) / base.total * 100) : null;
  const resposta = ps.length > 1 && base.total > 0 && atual.total <= base.total / 2, remissao = atual.total < 5;
  const barra = t => `<table class="pb"><tr><td style="width:${Math.max(t, 0.4) / 27 * 100}%;background:${PHQ_TONS[phqFaixa_(t)]}">&nbsp;</td><td style="background:#F1EFE8">&nbsp;</td></tr></table>`;
  const ponto = v => v == null ? "—" : ["○", "◔", "◑", "●"][v];
  return `<h2>Resultado de hoje</h2>
  <table class="sint"><tr>
   <td><b>PHQ-9</b><span style="font-size:20pt">${atual.total}</span> / 27<br>sintomas ${PHQ_FAIXAS[phqFaixa_(atual.total)][2]}</td>
   <td><b>Em relação ao início</b>${ps.length > 1 ? `Basal ${base.total} (${data(base.quando)}) → ${atual.total}<br>${varPct <= 0 ? "Queda" : "Aumento"} de ${Math.abs(varPct)}%` : "Primeiro PHQ-9: ponto de partida"}</td>
   <td><b>Resposta · remissão</b>Resposta (queda ≥ 50%): ${ps.length > 1 ? (resposta ? "<b style='display:inline;font-size:9pt;color:#2F5A48'>sim</b>" : "não") : "—"}<br>Remissão (PHQ-9 &lt; 5): ${remissao ? "<b style='display:inline;font-size:9pt;color:#2F5A48'>sim</b>" : "não"}<br><span class="mini">Dificuldade funcional: ${e(atual.func || "—")}</span></td></tr></table>
  <h2>Evolução do PHQ-9</h2>
  <div class="leg" style="margin:0 0 4px">Cada linha é um questionário. A cor da barra acompanha a faixa: ${PHQ_FAIXAS.map((f, i) => `<span style="display:inline-block;width:9px;height:9px;background:${PHQ_TONS[i]};border-radius:2px"></span> ${f[0]}–${f[1]} ${f[2]}`).join(" · ")}.</div>
  <table class="graf">${ps.map((r, i) => `<tr><td class="dl">${data(r.quando)}</td><td>${barra(r.total)}</td><td class="dv"><b>${r.total}</b></td><td class="dk">${i ? (r.total - ps[i - 1].total > 0 ? "+" : "") + (r.total - ps[i - 1].total) + " desde o anterior" : "basal"}${r.itens[8] > 0 ? " · <b style='display:inline;color:#83462D'>item 9: " + r.itens[8] + "</b>" : ""}</td></tr>`).join("")}</table>
  <h2>Itens</h2>
  <table class="r"><tr><th style="text-align:left">Sintoma</th>${ps.length > 2 ? "<th>Basal</th>" : ""}${ant ? "<th>Anterior</th>" : ""}<th>Hoje</th></tr>
  ${PHQ_CURTO.map((q, i) => `<tr class="${i === 8 && atual.itens[8] > 0 ? "rk" : atual.itens[i] >= 2 ? "g3" : ""}"><td>${i + 1}. ${e(q)}</td>${ps.length > 2 ? `<td class="c">${base.itens[i]}</td>` : ""}${ant ? `<td class="c">${ant.itens[i]}</td>` : ""}<td class="c"><b>${atual.itens[i]}</b></td></tr>`).join("")}</table>
  <div class="leg">0 = nenhuma vez · 1 = vários dias · 2 = mais da metade dos dias · 3 = quase todos os dias.</div>
  <h2>${ant ? "Check-ins desde o PHQ-9 anterior" : "Check-ins"} e hoje</h2>
  <table class="r"><tr><th style="text-align:left">Data</th><th>Humor</th><th style="text-align:left">Sono</th><th style="text-align:left">Remédios</th><th style="text-align:left">Efeitos colaterais</th><th>Morte/ferir</th></tr>
  ${desde.concat([atual]).map(r => `<tr class="${r.morte === "SIM" ? "rk" : ""}"><td>${data(r.quando)}${r.tipo === "PHQ-9" ? " (PHQ)" : ""}</td><td class="c">${e(r.humor || "—")}</td><td>${e(r.sono || "—")}</td><td>${e(r.remedio || "—")}</td><td>${e(r.efeitos || "—")}</td><td class="c">${r.tipo === "PHQ-9" ? "item 9: " + r.itens[8] : e(r.morte || "—")}</td></tr>`).join("")}</table>
  ${desde.some(r => r.nota) ? `<h3>Notas dos check-ins</h3>${desde.filter(r => r.nota).map(r => `<p><b style="display:inline">${data(r.quando)}:</b> ${e(r.nota)}</p>`).join("")}` : ""}`;
}

function montarHtml(d) {
  const e = esc_, curto = q => q.split(/,| \(/)[0];
  const abertas = d.secoes.filter(s => s.titulo !== "Identificação" && s !== (d.rastreio && d.secoes.find(x => x.titulo === d.rastreio.titulo)) &&
    !(d.impacto && s.titulo === d.impacto.titulo) && s.titulo !== "Motivação" && !/^Resumo para o médico/.test(s.titulo) &&
    (d.roteiro !== "PHQ-9" || s.titulo === "Observações do paciente"));
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
  table.graf{width:100%;border-collapse:collapse;margin:2px 0 4px}
  table.graf td{padding:3px 4px;vertical-align:middle}
  td.dl{width:62px;font-size:9pt;color:#5f5d4c}
  td.dv{width:40px;font-size:10pt;text-align:right}
  td.dk{width:150px;font-size:8.5pt;color:#5f5d4c}
  table.seg{width:100%;table-layout:fixed;border-collapse:separate;border-spacing:2px 0}
  table.seg td{width:11.1%;height:15px;padding:0;border-radius:2px;font-size:1pt;line-height:15px}
  table.dv-leg td{font-size:8.5pt;color:#2B2A22;padding-right:16px}
  table.dv-leg span{display:inline-block;width:10px;height:10px;border-radius:2px;vertical-align:-1px}
  table.dv-leg i{display:inline-block;width:0;height:11px;border-left:2.5px solid #2B2A22;vertical-align:-1px;margin:0 3px}
  table.sint{width:100%;border-collapse:separate;border-spacing:6px 0;margin:8px -6px 0}
  table.sint td{width:33%;vertical-align:top;background:#F7F5EF;border:1px solid #e2dccb;border-radius:6px;padding:6px 8px;font-size:9pt}
  table.sint b{display:block;font-size:7.5pt;text-transform:uppercase;letter-spacing:.04em;color:#5f5d4c;margin-bottom:2px}
  .mini{font-size:8pt;color:#5f5d4c}
  table.pb{width:100%;border-collapse:collapse;table-layout:fixed}
  table.pb td{height:14px;padding:0;font-size:1pt;line-height:14px}
  table.mapa{border-collapse:collapse;margin-top:4px}
  table.mapa th{font-size:8pt;color:#5f5d4c;font-weight:normal;padding:1px 6px}
  table.mapa td.c{text-align:center;font-size:12pt;padding:0 6px;line-height:1.1}
  .nota{font-size:8pt;color:#5f5d4c;border-top:1px solid #e2dccb;margin-top:12px;padding-top:5px}
  </style></head><body>
  <div class="cab"><h1>${{ "DIVA-5": "DIVA-5 · TDAH em adultos", "PHQ-9": "Acompanhamento da depressão" }[d.roteiro] || "Pré-consulta"} <span class="tag">${e((d.roteiro || "").toUpperCase())}</span></h1>
  <div class="dr">Dr. Luis Felipe Jardim · Psiquiatra<br>CRM-SC 12767 · RQE 22739<br>Recebida em ${e(d.enviado)}</div></div>
  <div class="id"><div class="nome">${e(d.nome)}</div><table><tr>
   <td><b>CPF</b>${e(d.cpfFmt)}</td><td><b>Nascimento · idade</b>${e(d.nasc)}${d.idade != null ? " · " + d.idade + " anos" : ""}</td><td><b>Celular</b>${e(d.tel)}</td></tr><tr>
   <td><b>${d.roteiro === "DIVA-5" ? "Ajuda com a infância" : "Respondido por"}</b>${e(d.roteiro === "DIVA-5" ? (d.ident["Quem ajudou com a infância"] || "—") : d.respondente)}${d.cuidador ? "<br>" + e(d.cuidador) : ""}</td>
   <td><b>Cadastro V Saúde</b>${d.vsaude ? (d.vsaude.numero ? "Paciente nº " + e(d.vsaude.numero) + " · l" : "L") + "ocalizado pelo " + e(d.vsaude.via || "CPF") : "—"}</td>
   <td><b>Tempo de preenchimento</b>${e(d.tempo || "—")}</td></tr></table></div>
  ${d.risco || d.riscoItens.length ? `<div class="alerta"><b>⚠ Sinal de risco.</b> ${d.riscoItens.map(it => `Rastreio: “${e(it.q)}” = <b>${FREQ[it.v]}</b>.`).join(" ")}
    ${frases.length ? " No relato: " + frases.map(f => "“" + e(f) + "”").join(" ") : ""} A orientação do CVV 188 / SAMU 192 foi exibida na tela.</div>` : ""}
  ${d.roteiro === "DIVA-5" || d.roteiro === "PHQ-9" ? "" : `<h2>Visão rápida</h2>
  <table class="vr" style="width:100%;border-collapse:separate;border-spacing:6px 0"><tr>
   <td><b>Quanto atrapalha hoje</b>${d.impacto ? `<span class="big">${e(d.impacto.nota || "—")}</span> / 10<br>${e(d.impacto.areas.split(", ").join(" · "))}` : "—"}</td>
   <td><b>Motivação</b>Importância de resolver agora: <strong>${e(d.motivacao.imp || "—")}</strong>/10<br>Confiança de mudar com ajuda: <strong>${e(d.motivacao.conf || "—")}</strong>/10</td>
   <td><b>Rastreio: moderado ou grave</b>${destaques.length ? destaques.map(it => "• " + e(curto(it.q))).join("<br>") : "Nenhum item"}</td></tr></table>`}
  ${d.rastreio ? `<h2>Rastreio · ${e(d.rastreio.titulo.toLowerCase())}</h2>
  <table class="r"><tr><th></th><th style="text-align:left">Item</th>${FREQ.map(f => `<th>${f}</th>`).join("")}</tr>
  ${d.rastreio.itens.map((it, i) => `<tr class="${it.risco && it.v > 0 ? "rk" : it.v === 4 ? "g4" : it.v === 3 ? "g3" : ""}"><td style="color:#5f5d4c;text-align:right">${i + 1}</td><td>${e(it.q)}</td>${FREQ.map((_, j) => `<td class="c">${it.v === j ? "●" : ""}</td>`).join("")}</tr>`).join("")}</table>
  <div class="leg">Muito leve = 1–2 dias · Leve = vários dias · Moderado = mais da metade dos dias · Grave = quase todos os dias. Itens sem marcação: não respondidos.</div>` : ""}
  ${d.roteiro === "DIVA-5" ? divaHtml_(d) : d.roteiro === "PHQ-9" ? phqHtml_(d) : ""}
  ${d.roteiro === "PHQ-9" ? (abertas[0] && abertas[0].texto ? `<h2>Observações do paciente</h2><div class="leg" style="margin:0 0 4px">Texto digitado ou ditado pelo paciente.</div>${abertas[0].texto.split("\n").map(l => `<p>${e(l)}</p>`).join("")}` : "") : `
  <h2>${d.roteiro === "DIVA-5" ? "Respostas, item por item" : "Respostas abertas"}</h2>
  <div class="leg" style="margin:0 0 4px">Texto ditado ou digitado por quem respondeu e revisado antes do envio. A oralidade foi preservada.</div>
  ${abertas.map(s => `<h3>${e(s.titulo)}</h3>${s.texto ? s.texto.split("\n").map(l => `<p>${e(l)}</p>`).join("") : '<p class="vazio">Não respondido.</p>'}`).join("")}`}
  <div class="nota">Documento gerado automaticamente a partir do questionário on-line e anexado ao prontuário. Conteúdo sigiloso, protegido pelo sigilo médico e pela LGPD (Lei 13.709/2018).</div>
  </body></html>`;
}

if (typeof module !== "undefined") module.exports = { interpretar, montarHtml, localizar_, norm_, mesmoNome_, registroHumor_ }; // testes locais
