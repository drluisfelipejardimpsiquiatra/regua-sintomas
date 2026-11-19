// Sincroniza a agenda do vSaúde com o Mural de Receitas (somente leitura na API).
//
// Uso:  NODE_USE_ENV_PROXY=1 node mural-receitas/sync-vsaude.mjs <saida.json> [diretorio.json]
// O segundo arquivo (opcional) é a lista de pacientes para o cadastro do Monitor de Peso:
// { atualizadoEm, lista: "<JSON [{id, n: nome, t: telefone, u: última consulta, p: próxima}]>" } — sem CPF.
//
// A chave da API é injetada pelo proxy do ambiente na nuvem (cabeçalho VSAUDE-API-KEY);
// localmente, defina VSAUDE_API_KEY. Nada aqui grava no vSaúde.
// Saída: { atualizadoEm, linhas: "<JSON>" } — um paciente por linha, sem CPF, telefone ou dados clínicos.

import { writeFileSync } from "node:fs";

const BASE = "https://public-api.vsaude.com.br/api/services/app";
const MEDICO = "8ff03fa9-0ef3-4d11-a6dc-b1dd00878475";
const UNIDADES = { 5398: "Torres", 4177: "Sombrio", 4176: "Praia Grande" };
// Procedimentos em que há avaliação médica (e portanto receita). Fora: bloqueio, palestra, reunião.
const CONSULTA = /consulta|retorno|avalia|escetamina|ect/i;
const ATENDIDO = new Set([81]); // 81 = atendido/finalizado (docs/vsaude-api.md do projeto de triagem)

async function call(path, body) {
  const headers = { "Content-Type": "application/json" };
  if (process.env.VSAUDE_API_KEY) headers["VSAUDE-API-KEY"] = process.env.VSAUDE_API_KEY;
  for (let tentativa = 1; ; tentativa++) {
    const res = await fetch(`${BASE}/${path}`, { method: "POST", headers, body: JSON.stringify(body) });
    const txt = await res.text();
    let d; try { d = JSON.parse(txt); } catch { d = { raw: txt.slice(0, 200) }; }
    if (res.ok && d.success !== false) return d.result ?? d;
    if (tentativa >= 3 || res.status === 401) throw new Error(`${path}: HTTP ${res.status} ${d?.error?.message || d.raw || ""}`);
    await new Promise((r) => setTimeout(r, 3000 * tentativa));
  }
}
const items = (r) => (Array.isArray(r) ? r : r?.items ?? []);
const diaBR = (iso) => new Date(iso).toLocaleDateString("en-CA", { timeZone: "America/Sao_Paulo" });
const horaBR = (iso) => new Date(iso).toLocaleTimeString("pt-BR", { timeZone: "America/Sao_Paulo", hour: "2-digit", minute: "2-digit" });

const agora = new Date();
const hoje = diaBR(agora.toISOString());
const desde = diaBR(new Date(agora.getTime() - 150 * 864e5).toISOString());

// GetAll devolve ~6 meses a partir de `from`: duas janelas cobrem 150 dias para trás e ~6 meses à frente.
const janelas = await Promise.all([desde, hoje].map((from) => call("ScheduleService/GetAll", { professionals: [MEDICO], from, range: 4 })));
const porId = new Map();
for (const j of janelas) for (const a of items(j)) porId.set(a.id, a);

const pac = new Map();
for (const a of porId.values()) {
  const proc = (a.procedure?.name || a.title || "").trim();
  if (!a.patient?.id || !CONSULTA.test(proc)) continue;
  const dia = diaBR(a.date);
  const p = pac.get(a.patient.id) || { pid: a.patient.id, nome: (a.patient.name || "").replace(/\s+/g, " ").trim(), ultima: "", ultimaProc: "", ultimaUnidade: "", proxima: "", proximaHora: "", proximaProc: "", proximaUnidade: "", atendidas: 0 };
  if (dia <= hoje && ATENDIDO.has(a.status)) {
    p.atendidas++;
    if (dia > p.ultima) { p.ultima = dia; p.ultimaProc = proc; p.ultimaUnidade = UNIDADES[a.careUnit?.id] || ""; }
  } else if (dia > hoje && (!p.proxima || dia < p.proxima)) {
    p.proxima = dia; p.proximaHora = horaBR(a.date); p.proximaProc = proc; p.proximaUnidade = UNIDADES[a.careUnit?.id] || "";
  }
  pac.set(a.patient.id, p);
}
if (process.argv[3]) {
  const dir = new Map();
  for (const a of porId.values()) {
    if (!a.patient?.id || !a.patient.name) continue;
    const dia = diaBR(a.date);
    const d = dir.get(a.patient.id) || { id: a.patient.id, n: a.patient.name.replace(/\s+/g, " ").trim(), t: a.patient.phoneNumber || "", u: "", p: "" };
    if (!d.t && a.patient.phoneNumber) d.t = a.patient.phoneNumber;
    if (dia <= hoje) { if (dia > d.u) d.u = dia; } else if (!d.p || dia < d.p) d.p = dia;
    dir.set(a.patient.id, d);
  }
  const lista = [...dir.values()].sort((a, b) => a.n.localeCompare(b.n, "pt-BR"));
  writeFileSync(process.argv[3], JSON.stringify({ atualizadoEm: agora.toISOString(), total: lista.length, lista: JSON.stringify(lista) }));
  console.log(`diretório: ${lista.length} pacientes`);
}
const linhas = [...pac.values()].filter((p) => p.ultima);
const saida = { atualizadoEm: agora.toISOString(), agendamentos: porId.size, pacientes: linhas.length, linhas: JSON.stringify(linhas) };
writeFileSync(process.argv[2] || "vsaude-mural.json", JSON.stringify(saida));
console.log(`ok: ${porId.size} agendamentos, ${linhas.length} pacientes com consulta atendida nos últimos 150 dias`);
