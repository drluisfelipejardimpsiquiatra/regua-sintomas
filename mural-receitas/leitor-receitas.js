// Leitor de receitas do vSaúde (somente leitura). Rodar com javascript_tool numa aba logada em https://app.vsaude.com.br
// Variável DIAS: quantos dias para trás olhar na agenda (padrão 3).
const DIAS = typeof window.__DIAS === 'number' ? window.__DIAS : 3;
const H = {Authorization: 'Bearer ' + decodeURIComponent((document.cookie.match(/Abp\.AuthToken=([^;]+)/) || [])[1] || '')};
const diaBR = (iso) => new Date(iso).toLocaleDateString('en-CA', {timeZone: 'America/Sao_Paulo'});
const toD = s => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)), toS = ms => new Date(ms).toISOString().slice(0, 10), add = (s, d) => toS(toD(s) + d * 864e5);
const strip = s => (s || '').replace(/&nbsp;/g, ' ').replace(/<[^>]+>/g, ' ').replace(/[‑–]/g, '-').replace(/\s+/g, ' ').trim();
const PSI = /vortiox|vognus|vorxe|evortia|voextor|brintellix|trazod|azod|donaren|venlaf|venlax|venlift|efexor|desvenl|desve\b|desduo|pristiq|zetron|bupropi|bup\b|bup xl|bupium|wellbutrin|sertral|assert|tolrest|zoloft|escitalo|reconter|exodus|lexapro|\besc\b|citalo|citta|città|fluoxe|daforin|prozac|paroxe|sincro|roxetin|pondera|aropax|mirtaz|razapina|remeron|menelat|duloxe|cymbi|velija|amitript|nortript|pamelor|clomipr|anafranil|agomel|valdoxan|vilazod|aymee|arpejo|amato|quetia|\bquet\b|quetipin|quetros|quepsia|seroquel|risper|\briss\b|olanza|\bzap\b|zyprexa|aripip|aristab|toarip|abilify|brexpip|rexulti|lurasid|latuda|ziprasid|geodon|cariprazin|reagila|clozap|leponex|haloper|haldol|levomeprom|neozine|levozine|clorprom|amplictil|periciaz|neuleptil|sulpir|lamotrig|lamitor|neural|valpro|depakote|depakene|divalc|torval|carbamaz|tegretol|oxcarbaz|trileptal|topiram|topamax|gabapent|gaba er|neurontin|pregabal|lyrica|prebictal|insit|litio|lítio|carbolit|carlit|clonaz|rivotril|alpraz|frontal|lorazep|lorax|diazep|bromazep|lexotan|zolpid|stilnox|patz|eszopic|prysma|metilfenid|ritalina|concerta|consiv|lisdexanf|venvanse|juneve|lisdev|modafin|stavigile|naltrex|uninaltrex|revia|dissulfir|antabus|buspir|stima|ansitec|memant|donepez|rivastig|galantam|atomox|atentah|motraz|sonic|lyberdia|lidexor|afetus|dorene|limiar|alpes|inseris|konduz|vendexla|glya|lubip|atip|mobale|dep\b|don\b|vod\b|eudok|hezo|seth|lurasidona/i;
const num = s => { s = s.replace(',', '.').trim(); let m; if ((m = s.match(/^(\d+)\s*\+\s*1\/2$/))) return +m[1] + .5; if ((m = s.match(/^(\d+)\/(\d+)$/))) return m[1] / m[2]; return parseFloat(s); };
const doseDia = (t) => { t = (t || '').toLowerCase(); const c = t.slice(1).search(/\d+\s*-\s*[a-zà-ú]|\d+\s*mg\b/); if (c >= 0) t = t.slice(0, c + 1); const k = t.search(/\b(após|apos|depois)\b/); if (k >= 0) t = t.slice(k);
  const re = /(\d+\s*\+\s*1\/2|1\/2|\d+(?:[.,]\d+)?)\s*(?:cp|cps|comp\w*|caps|c[aá]psulas?|dr[aá]geas?)\b/g; let m, tot = 0, n = 0; while ((m = re.exec(t))) { const v = num(m[1]); if (v > 8) continue; tot += v; n++; }
  if (!n) return null; let mult = 1; if (/12\s*\/\s*12|2\s*x\s*(ao|por)\s*dia|2 vezes/.test(t)) mult = 2; else if (/8\s*\/\s*8|3\s*x\s*(ao|por)\s*dia|3 vezes|café.*almoço.*jantar/.test(t)) mult = 3;
  if (n > 1) mult = 1; return tot * mult; };
const qtd = (s) => { let m = s.match(/(\d+)\s*(cps|cp|comp\w*|c[aá]ps\w*|dr[aá]geas?)\b/i); if (m && +m[1] >= 10) return {n: +m[1], cx: 0}; m = s.match(/(\d+)\s*(cxs?|caixas?)\b/i); if (m) return {n: +m[1] * 30, cx: +m[1]}; return null; };
const VERBO = /\b(tomar|aplicar|usar|pingar|dissolver|colocar|ingerir)\b/i;
const rxDe = async (pid) => { const j = await (await fetch('/api/services/app/PrescriptionService/GetPrescriptionsByPatientId?patientId=' + pid, {headers: H})).json(); const a = j.result; return (Array.isArray(a) ? a : (a.items || [])).filter(x => !x.isDeleted); };
const curto = (n) => n.replace(/^(Hemifumarato|Oxalato|Cloridrato|Succinato|Bromidrato) De /i, '').split(' ').slice(0, 2).join(' ').slice(0, 14).replace(/,/g, '.');
const proc = async (p) => {
  let a; for (let t = 0; t < 3; t++) { try { a = await rxDe(p.pid); break; } catch (e) { await new Promise(r => setTimeout(r, 1500)); } }
  if (!a) return null;
  const da = a.filter(x => diaBR(x.creationTime) >= add(p.ultima, -1)); if (!da.length) return null;
  const itens = {}; let cx = false;
  const put = (k, e) => { if (!itens[k] || itens[k].de < e.de) itens[k] = e; };
  for (const x of da) { const dt = diaBR(x.creationTime);
    for (const m of x.medications) { const raw = strip(m.drugDescriptionAndDosage); const ins = strip(m.usageInstruction); if (!PSI.test(raw)) continue;
      const pt = raw.split('·').map(s => s.trim()); const nome = (pt.length > 2 ? pt[0] + ' ' + ((pt[2] || '').match(/[\d,]+\s*MG/i) || [''])[0] : raw.replace(/\s\d+\s*(cps|cp|comp\w*|c[aá]ps\w*|cxs?|caixas?)\b.*$/i, '')).toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase()).trim();
      const dd = doseDia(ins); let dias = null, star = '';
      const qn = m.medicationCatalogId ? null : qtd(raw);
      if (qn && dd) { dias = Math.round(qn.n / dd); if (qn.cx) star = '*'; }
      else if (m.period && m.period.unit === 1) { dias = m.period.value; const mc = ins.match(/(\d+)\s*(cxs?|caixas?)\b/i); if (mc && dd && Math.abs(Math.round(+mc[1] * 30 / dd) - dias) > 7) cx = true; }
      else { dias = 60; star = '*'; }
      put(nome.split(' ')[0].toLowerCase(), {n: curto(nome), d: Math.min(dias, 60), de: dt, s: star, t: /cessar|suspender|at[eé] parar/i.test(ins)}); }
    const txt = strip(x.content);
    if (txt) for (const s of txt.split(/(?:^|\s)\d+\s*-\s*(?=[A-Za-zÀ-ú])/).map(s => s.trim()).filter(Boolean)) {
      if (!PSI.test(s)) continue;
      const iv = s.search(VERBO); const cab = iv > 0 ? s.slice(0, iv) : s; const pos = iv > 0 ? s.slice(iv) : '';
      const qn = qtd(cab); const nome = cab.replace(/\s\d+\s*(cps|cp|comp\w*|c[aá]ps\w*|cxs?|caixas?)\b.*$/i, '').trim(); const dd = doseDia(pos);
      let dias = null, star = ''; if (qn && dd) dias = Math.round(qn.n / dd); if (!dias) { dias = 60; star = '*'; } if (qn && qn.cx) star = '*';
      put(nome.split(' ')[0].toLowerCase(), {n: curto(nome), d: Math.min(dias, 60), de: dt, s: star, t: /cessar|suspender/i.test(s)}); }
  }
  const its = Object.values(itens); if (!its.length) return null;
  const base = its.filter(i => !i.t); const ate = (base.length ? base : its).map(i => add(i.de, i.d)).sort()[0];
  return [p.pid, p.ultima, ate, its.map(i => i.n + ' ' + i.d + i.s).join(','), cx ? 'cx' : ''].join('\t');
};
const hoje = diaBR(new Date().toISOString()), desde = diaBR(new Date(Date.now() - DIAS * 864e5).toISOString());
const ag = await (await fetch('/api/services/app/ScheduleService/GetAll', {method: 'POST', headers: {...H, 'Content-Type': 'application/json'}, body: JSON.stringify({professionals: ['8ff03fa9-0ef3-4d11-a6dc-b1dd00878475'], from: desde, range: 1})})).json();
const lista = Array.isArray(ag.result) ? ag.result : (ag.result.items || []);
const pac = new Map();
for (const a of lista) { const proc = ((a.procedure && a.procedure.name) || a.title || ''); if (!a.patient || !a.patient.id || !/consulta|retorno|avalia|escetamina|ect/i.test(proc)) continue; const d = diaBR(a.date); if (d < desde || d > hoje || a.status !== 81) continue; const p = pac.get(a.patient.id) || {pid: a.patient.id, ultima: ''}; if (d > p.ultima) p.ultima = d; pac.set(a.patient.id, p); }
const out = []; const fila = [...pac.values()];
await Promise.all(Array.from({length: 4}, async () => { while (fila.length) { const r = await proc(fila.shift()); if (r) out.push(r); } }));
window.__linhas = out.join('\n');
const partes = []; for (let i = 0; i < window.__linhas.length; i += 900) partes.push(window.__linhas.slice(i, i + 900));
({pacientes: pac.size, lidos: out.length, partes});
