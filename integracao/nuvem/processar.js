#!/usr/bin/env node
/**
 * Integração V Saúde rodando na nuvem (sem instalar nada na planilha).
 * ---------------------------------------------------------------
 * Usa o MESMO programa da planilha (../vsaude-preconsulta.gs), com estes adaptadores:
 *   - planilha  → arquivo JSON com as linhas lidas pelo conector do Google Sheets;
 *   - V Saúde   → API pública (a chave é injetada pelo ambiente);
 *   - PDF       → gerado pelo Chromium a partir do mesmo HTML;
 *   - e-mail    → devolvido em JSON para a sessão enviar pelo Gmail.
 *
 * Uso: node processar.js linhas.json saida.json [--simular]
 *   linhas.json: { "inicio": "2026-11-25", "linhas": [[carimbo, paciente, respostas, status], ...] }  (linha 1 = cabeçalho)
 *   saida.json : { "status": { "<nº da linha>": "texto" }, "emails": [{ assunto, texto }], "anexos": [...] }
 * Só processa linhas a partir de "inicio" e que ainda não tenham status ✅ ou ⛔.
 */
const fs = require("fs"), vm = require("vm"), path = require("path"), os = require("os"), cp = require("child_process");
const [, , ENTRADA, SAIDA, ...flags] = process.argv;
const SIMULAR = flags.includes("--simular");
const BASE = "https://public-api.vsaude.com.br/api/services/app";
const dados = JSON.parse(fs.readFileSync(ENTRADA, "utf8"));
const linhas = dados.linhas.map((l) => [l[0], l[1] || "", l[2] || "", l[3] || ""]);
const COL_STATUS = 4;
if (linhas[0][COL_STATUS - 1] !== "Status V Saúde") linhas[0][COL_STATUS - 1] = "Status V Saúde";
const saida = { status: {}, emails: [], anexos: [] };
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "vsaude-"));

// "25/11/2026 18:43:10" (formato da planilha) → Date no fuso de Brasília
const paraData = (v) => {
  if (v instanceof Date) return v;
  const m = String(v).match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/);
  return m ? new Date(`${m[3]}-${m[2]}-${m[1]}T${m[4]}:${m[5]}:${m[6] || "00"}-03:00`) : new Date(v);
};
const fmt = (d, tz, f) => {
  const p = Object.fromEntries(new Intl.DateTimeFormat("en-GB", { timeZone: "America/Sao_Paulo", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" })
    .formatToParts(d).map((x) => [x.type, x.value]));
  return f.replace("dd", p.day).replace("MM", p.month).replace("yyyy", p.year).replace("yy", p.year.slice(2)).replace("HH", p.hour).replace("mm", p.minute);
};
const curl = (args) => cp.execFileSync("curl", ["-s", "--max-time", "60", ...args], { maxBuffer: 50 * 1024 * 1024 }).toString();

const planilha = {
  getLastRow: () => linhas.length, getLastColumn: () => COL_STATUS,
  getRange: (r, c, nr, nc) => {
    const ler = () => linhas.slice(r - 1, r - 1 + (nr || 1)).map((l) => l.slice(c - 1, c - 1 + (nc || 1)).map((v, j) => (c + j === 1 && r > 1 ? paraData(v) : v)));
    const rg = {
      getValues: ler, getValue: () => ler()[0][0],
      setValue: (v) => { linhas[r - 1][c - 1] = v; if (r > 1 && c === COL_STATUS) saida.status[r] = v; return rg; },
      setFontWeight: () => rg };
    return rg;
  } };
const ctx = {
  console, URL, encodeURIComponent, Intl,
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheets: () => [planilha], getUrl: () => "https://docs.google.com/spreadsheets/d/" + (dados.planilhaId || "") }) },
  Session: { getEffectiveUser: () => ({ getEmail: () => "medico" }) },
  MailApp: { sendEmail: (para, assunto, texto) => saida.emails.push({ assunto, texto }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => "injetada-pelo-ambiente" }) },
  Utilities: { formatDate: fmt },
  HtmlService: { createHtmlOutput: (html) => ({ getAs: () => ({ setName: (nome) => {
    const h = path.join(tmp, "doc.html"), pdf = path.join(tmp, Date.now() + ".pdf");
    fs.writeFileSync(h, html);
    cp.execFileSync(process.execPath, [path.join(__dirname, "pdf.js"), h, pdf]);
    return { nome, pdf };
  } }) }) },
  UrlFetchApp: { fetch: (url, op) => {
    const out = curl(["-X", (op.method || "GET").toUpperCase(), url, ...(op.payload && typeof op.payload === "string" ? ["-H", "Content-Type: application/json", "--data-binary", op.payload] : [])]);
    return { getResponseCode: () => 200, getContentText: () => out };
  } } };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "vsaude-preconsulta.gs"), "utf8") + "\n;this.processarLinha_ = processarLinha_;", ctx, { filename: "vsaude-preconsulta.gs" });
// Envio do PDF: multipart pelo curl (no Apps Script é o UrlFetchApp com o blob).
vm.runInContext("enviarArquivo_ = (id, nome, blob) => __enviar(id, nome, blob);", Object.assign(ctx, {
  __enviar: (id, nome, blob) => {
    saida.anexos.push({ paciente: id, nome, simulado: SIMULAR });
    if (SIMULAR) { fs.copyFileSync(blob.pdf, path.join(path.dirname(path.resolve(SAIDA)), nome.replace(/[\/\\]/g, "-"))); return {}; }
    const q = `?parent=00000000-0000-0000-0000-000000000000&ownerPatient=${id}&fileName=${encodeURIComponent(nome)}`;
    const r = JSON.parse(curl(["-X", "POST", BASE + "/FilesService/Upload" + q, "-F", `file=@${blob.pdf};type=application/pdf`]) || "{}");
    if (!r.success || !r.result || !r.result.uploadedFile) throw new Error("V Saúde não confirmou o anexo: " + JSON.stringify(r).slice(0, 150));
    return r.result;
  } }));

const inicio = dados.inicio ? new Date(dados.inicio + "T00:00:00-03:00") : new Date(0);
for (let r = 2; r <= linhas.length; r++) {
  const [carimbo, cab, , st] = linhas[r - 1];
  if (!cab || String(st).startsWith("✅") || String(st).startsWith("⛔")) continue;
  if (paraData(carimbo) < inicio) continue;
  ctx.processarLinha_(r, String(st).startsWith("⚠"));
}
fs.writeFileSync(SAIDA, JSON.stringify(saida, null, 2));
console.log(`linhas: ${linhas.length - 1} · status novos: ${Object.keys(saida.status).length} · anexos: ${saida.anexos.length}${SIMULAR ? " (simulados)" : ""} · e-mails: ${saida.emails.length}`);
