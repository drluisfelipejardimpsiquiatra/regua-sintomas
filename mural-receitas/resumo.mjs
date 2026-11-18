// Resumo do mural para a notificação da rotina.
// Uso: node mural-receitas/resumo.mjs <vsaude.json> [dir-fichas] [dir-receitas]
// Os diretórios são os gerados por ArtifactData list com out_dir (um JSON por documento).
// Reaproveita o cálculo do painel (mural-receitas/painel.html) para não divergir dele.
import { readFileSync, readdirSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, "painel.html"), "utf8");
const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)][0][1];
const C = {};
globalThis.dash = { calc: (id, o) => { C[id] = o.fn; } };
new Function(script)();

const lerDir = (d) => {
  if (!d || !existsSync(d)) return [];
  const out = [];
  const walk = (p) => { for (const f of readdirSync(p, { withFileTypes: true })) { const q = join(p, f.name); if (f.isDirectory()) walk(q); else if (f.name.endsWith(".json")) { const o = JSON.parse(readFileSync(q, "utf8")); const doc = o.data || o; out.push({ id: o.id || f.name.replace(/\.json$/, ""), ...doc }); } } };
  walk(d); return out;
};
const vs = JSON.parse(readFileSync(process.argv[2], "utf8"));
const mural = C.mural([vs], lerDir(process.argv[4]), lerDir(process.argv[3]));
const r = C.resumo(mural)[0];
const pad = (n) => String(n).padStart(2, "0");
const n = new Date(); const hoje = new Date(n.toLocaleString("en-US", { timeZone: "America/Sao_Paulo" }));
const hojeS = `${hoje.getFullYear()}-${pad(hoje.getMonth() + 1)}-${pad(hoje.getDate())}`;
const novos = mural.filter((x) => x.situacao === "renovar" && x.alertaEm === hojeS);
const fmt = (s) => `${s.slice(8, 10)}/${s.slice(5, 7)}`;
console.log(`Renovar agora: ${r.agora} · próximos 21 dias: ${r.breve} · cobertos: ${r.cobertos} · sem próxima consulta: ${r.semProxima}`);
console.log(`Entraram hoje no mural: ${novos.length}`);
for (const x of novos) console.log(`- ${x.paciente}: receita até ${fmt(x.receitaAte)}, +${x.sugestaoMeses === 1 ? "1 mês" : "2 meses"}${x.proxima ? `, próxima consulta ${fmt(x.proxima)}` : ", sem próxima consulta"}`);
