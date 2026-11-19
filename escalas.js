// Escalas leigas de arrastar o dedo (inspiradas em instrumentos clínicos, não são as versões oficiais).
// Mesmo arquivo copiado no app privado (regua-app.html). Ao mudar aqui, mudar lá também.
const HORA = (v) => { const h = Math.floor(v) % 24, m = Math.round((v % 1) * 60); return String(h).padStart(2, "0") + "h" + (m ? String(m).padStart(2, "0") : ""); };
const ESCALAS = {
  regua: {
    nome: "Régua de sintomas", curto: "Depressão", periodo: "nos últimos dias", max: 90,
    itens: [
      ["tr", "Tristeza"], ["fp", "Falta de prazer", "nas coisas de que você gostava"], ["cm", "Crises de medo", "ou de pânico"],
      ["te", "Tensão / preocupação"], ["ir", "Irritabilidade / impaciência"], ["ap", "Aumento de apetite"],
      ["so", "Problemas de sono", "dificuldade para dormir, acordar à noite ou sono demais"], ["en", "Falta de energia", "cansaço, pouca disposição"],
      ["pm", "Pensamentos de morte", "ou de que seria melhor não estar vivo(a)", { risco: true }],
    ],
  },
  ansiedade: {
    nome: "Ansiedade no dia a dia", curto: "Ansiedade", periodo: "nos últimos dias", max: 80,
    itens: [
      ["a1", "Nervosismo", "ou sensação de estar no limite"], ["a2", "Preocupação que não consigo parar"],
      ["a3", "Preocupação com muitas coisas", "ao mesmo tempo"], ["a4", "Dificuldade para relaxar"],
      ["a5", "Inquietação", "difícil ficar parado(a)"], ["a6", "Irritação fácil"],
      ["a7", "Medo de que algo ruim vá acontecer"], ["a8", "Quanto isso atrapalhou meu dia a dia"],
    ],
  },
  panico: {
    nome: "Crises de pânico", curto: "Pânico", periodo: "na última semana", max: 50,
    itens: [
      ["p0", "Quantas crises fortes tive", "medo muito forte que chega de repente, com coração acelerado, falta de ar ou tontura", { min: 0, max: 7, fmt: (v) => (v >= 7 ? "7 ou mais" : String(v)), semCor: true, foraTotal: true, rotulos: ["nenhuma", "7 ou mais"] }],
      ["p1", "Força das crises"], ["p2", "Medo de ter uma nova crise"],
      ["p3", "Evitei lugares ou situações", "por medo de passar mal"], ["p4", "Sintomas no corpo", "coração acelerado, falta de ar, tontura"],
      ["p5", "Quanto isso atrapalhou meu dia a dia"],
    ],
  },
  toc: {
    nome: "Pensamentos repetitivos e manias", curto: "TOC", periodo: "na última semana", max: 80,
    intro: "Pensamentos que voltam sem você querer e coisas que sente que precisa fazer ou repetir (rituais, manias).",
    itens: [
      ["o1", "Tempo ocupado por pensamentos", "que voltam sem eu querer"], ["o2", "Quanto esses pensamentos", "atrapalham minha vida"],
      ["o3", "Angústia que esses pensamentos me causam"], ["o4", "Dificuldade de afastar esses pensamentos"],
      ["c1", "Tempo gasto com rituais ou manias", "lavar, conferir, arrumar, contar, repetir"], ["c2", "Quanto os rituais", "atrapalham minha vida"],
      ["c3", "Aflição se eu não fizer o ritual"], ["c4", "Dificuldade de deixar de fazer o ritual"],
    ],
  },
  sono: {
    nome: "Diário do sono", curto: "Sono", periodo: "sobre a noite que passou", max: null, diario: true,
    intro: "Responda de manhã, sobre a noite que passou. Leva 1 minuto.",
    itens: [
      ["deitar", "A que horas fui deitar", "", { min: 20, max: 28, step: 0.5, ini: 23, fmt: HORA, semCor: true, rotulos: ["20h", "4h"] }],
      ["lat", "Quanto tempo levei para pegar no sono", "", { min: 0, max: 180, step: 15, ini: 15, fmt: (v) => (v >= 180 ? "3 h ou mais" : v < 60 ? v + " min" : Math.floor(v / 60) + " h" + (v % 60 ? " " + (v % 60) + " min" : "")), rotulos: ["na hora", "3 h ou mais"], cor: (v) => (v <= 30 ? 0 : v <= 60 ? 5 : 9) }],
      ["desp", "Quantas vezes acordei durante a noite", "", { min: 0, max: 6, ini: 0, fmt: (v) => (v >= 6 ? "6 ou mais" : String(v)), rotulos: ["nenhuma", "6 ou mais"], cor: (v) => (v <= 1 ? 0 : v <= 3 ? 5 : 9) }],
      ["acordar", "A que horas acordei de vez", "", { min: 4, max: 12, step: 0.5, ini: 7, fmt: HORA, semCor: true, rotulos: ["4h", "12h"] }],
      ["horas", "Quantas horas dormi, mais ou menos", "", { min: 0, max: 12, step: 0.5, ini: 7, fmt: (v) => String(v).replace(".", ",") + " h", rotulos: ["0", "12 h"], cor: (v) => (v >= 7 ? 0 : v >= 5.5 ? 5 : 9) }],
      ["qual", "Qualidade do sono", "0 = péssima · 10 = ótima", { ini: 5, inverso: true, rotulos: ["péssima", "ótima"] }],
      ["desc", "Acordei descansado(a)", "0 = nada · 10 = totalmente", { ini: 5, inverso: true, rotulos: ["nada", "totalmente"] }],
      ["rem", "Usei remédio para dormir", "", { opcoes: ["Não", "Sim, o receitado", "Sim, outro"] }],
      ["coch", "Cochilei durante o dia de ontem", "", { opcoes: ["Não", "Menos de 30 min", "Mais de 30 min"] }],
    ],
  },
};
const ESC_ORDEM = ["regua", "ansiedade", "panico", "toc", "sono"];
const itemCfg = (it) => Object.assign({ min: 0, max: 10, step: 1, ini: 0 }, it[3] || {});
const totalEsc = (esc, v) => { const E = ESCALAS[esc]; if (!E || E.max == null) return null; return E.itens.reduce((s, it) => s + (itemCfg(it).foraTotal || itemCfg(it).opcoes ? 0 : (+v[it[0]] || 0)), 0); };
// nível de alerta de 0 a 10 de um item (para cor): maior = pior
const nivel = (it, v) => { const c = itemCfg(it); if (c.opcoes || c.semCor) return null; if (c.cor) return c.cor(v); const p = (v - c.min) / (c.max - c.min) * 10; return c.inverso ? 10 - p : p; };
