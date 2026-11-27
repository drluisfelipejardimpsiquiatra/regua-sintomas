#!/usr/bin/env node
// Gera Financeiro-completo.gs: Codigo.gs + as telas embutidas, para colar UM arquivo só no Apps Script.
// Uso: node montar-arquivo-unico.js   (rode de novo sempre que mudar Codigo.gs ou um .html)
const fs = require("fs"), path = require("path"), d = __dirname;
const telas = {}; ["Caixa", "Painel", "Estilo"].forEach((n) => { telas[n] = fs.readFileSync(path.join(d, n + ".html"), "utf8"); });
const saida = "// ARQUIVO GERADO por montar-arquivo-unico.js a partir de Codigo.gs, Caixa.html, Painel.html e Estilo.html.\n" +
  "// Cole este arquivo inteiro no Apps Script (não precisa criar os arquivos HTML).\n\n" +
  fs.readFileSync(path.join(d, "Codigo.gs"), "utf8") + "\n\n// ===================== TELAS EMBUTIDAS =====================\n\nconst ARQUIVOS_ = " + JSON.stringify(telas, null, 1) + ";\n";
fs.writeFileSync(path.join(d, "Financeiro-completo.gs"), saida);
console.log("Financeiro-completo.gs:", Math.round(saida.length / 1024), "KB");
