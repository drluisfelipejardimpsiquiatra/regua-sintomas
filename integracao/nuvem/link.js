#!/usr/bin/env node
/**
 * Gera o link pessoal de um paciente do V Saúde para os apps (Meu humor, DIVA-5, Y-BOCS/TOC).
 * Uso: node link.js <CPF ou nome> [humor|diva|toc]
 * Mostra os cadastros encontrados; com um único cadastro, imprime o código, os links e a linha para a aba "Links".
 * A linha da aba "Links" é gravada por quem chama (conector do Google Sheets).
 */
const cp = require("child_process"), crypto = require("crypto");
const [, , termo, app] = process.argv;
if (!termo) { console.error("Uso: node link.js <CPF ou nome> [humor|diva|toc]"); process.exit(1); }
const API = "https://public-api.vsaude.com.br/api/services/app/PatientService";
const req = (m, url) => JSON.parse(cp.execFileSync("curl", ["-s", "--max-time", "30", "-X", m, url]).toString());
const digitos = termo.replace(/\D/g, "");
const chave = digitos.length === 11 ? digitos.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, "$1.$2.$3-$4") : termo;
const vistos = {};
[chave, digitos.length === 11 ? digitos : null].filter(Boolean).forEach((k) => (req("POST", `${API}/Search?keyword=${encodeURIComponent(k)}`).result || [])
  .forEach((p) => { if (digitos.length !== 11 || String(p.personalIdentifier || "").replace(/\D/g, "") === digitos) vistos[p.id] = p; }));
const lista = Object.values(vistos).map((p) => {
  const g = req("GET", `${API}/Get?Id=${p.id}`).result || {};
  const b = String(g.birthday || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return { id: p.id, nome: p.name, cpf: p.personalIdentifier || "", nasc: b ? `${b[3]}/${b[2]}/${b[1]}` : "", tel: g.phone || p.phoneNumber || "" };
});
if (lista.length !== 1) { console.log(JSON.stringify({ encontrados: lista }, null, 2)); process.exit(lista.length ? 0 : 2); }
const p = lista[0], AB = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
const codigo = Array.from(crypto.randomBytes(8), (x) => AB[x % AB.length]).join("");
const primeiro = encodeURIComponent(p.nome.split(" ")[0]), base = "https://drluisfelipejardimpsiquiatra.github.io/consultorio/";
const links = { humor: `${base}humor.html?c=${codigo}&n=${primeiro}`, diva: `${base}diva.html?c=${codigo}&n=${primeiro}`, toc: `${base}toc.html?c=${codigo}&n=${primeiro}` };
console.log(JSON.stringify({ paciente: p, codigo, links: app ? { [app]: links[app] } : links,
  linhaAbaLinks: [codigo, p.id, p.nome, app === "diva" ? "DIVA-5" : app === "humor" ? "Meu humor" : app === "toc" ? "Y-BOCS (TOC)" : "Meu humor, DIVA-5 e Y-BOCS", new Date().toLocaleDateString("pt-BR", { timeZone: "America/Sao_Paulo" })] }, null, 2));
