// Converte o HTML do relatório em PDF A4 (usado por processar.js). Uso: node pdf.js entrada.html saida.pdf
const { chromium } = require("playwright");
(async () => {
  const [, , entrada, saida] = process.argv;
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto("file://" + require("path").resolve(entrada));
  await p.pdf({ path: saida, format: "A4", printBackground: true, margin: { top: "14mm", bottom: "16mm", left: "14mm", right: "14mm" } });
  await b.close();
})().catch((e) => { console.error(e); process.exit(1); });
