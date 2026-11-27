#!/usr/bin/env node
/**
 * Pré-consulta de ADULTO (formulário oficial "Questionário de pré-consulta – versão 10/2026").
 * ---------------------------------------------------------------
 * As respostas de adulto não chegam como texto (como criança, idoso e os apps): cada pergunta é uma coluna da
 * planilha "Questionário de pré-consulta: (respostas)". Este programa converte cada linha no mesmo texto organizado
 * ("[Adulto] nome · CPF ..." + seções "## ...") que o processar.js já transforma em PDF e anexa no V Saúde.
 *
 * Uso: node adulto.js bruto.json linhas.json
 *   bruto.json : { "inicio": "2026-11-25", "planilhaId": "...", "cabecalho": [linha 1 da aba, A..],
 *                  "linhas": [ { "linha": 427, "valores": [A..] } ],     (só as linhas a processar)
 *                  "status": { "427": "texto já gravado na coluna Status V Saúde, se houver" },
 *                  "links": [aba Links da planilha de recebimento, opcional] }
 *   linhas.json: entrada do processar.js; cada resposta no mesmo nº de linha da planilha de adulto.
 */
const fs = require("fs");
const [, , ENTRADA, SAIDA] = process.argv;
const dados = JSON.parse(fs.readFileSync(ENTRADA, "utf8"));
const norm = (s) => String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();

// Coluna → índice, pelo início do título da pergunta (títulos repetidos de versões antigas: vale o primeiro preenchido).
const cab = dados.cabecalho.map(norm);
const colunas = (inicio) => { const k = norm(inicio); return cab.map((h, i) => (h.startsWith(k) ? i : -1)).filter((i) => i >= 0); };
const valor = (vals, inicio) => { for (const i of colunas(inicio)) { const v = String(vals[i] == null ? "" : vals[i]).trim(); if (v) return v; } return ""; };

const FREQ = { "nada": "Nada", "muito leve": "Muito leve", "leve": "Leve", "moderado": "Moderado", "grave": "Grave" };
const curtoFreq = (v) => { const k = norm(v).replace(/\s*\(.*$/, ""); return FREQ[k] || (v ? v : "—"); };

// [rótulo no PDF, início do título da coluna]
const SECOES = [
  ["Identificação", [
    ["Nome completo", "Identificação (Nome do Paciente)"], ["CPF", "CPF"], ["Data de nascimento", "Data de nascimento"],
    ["Celular com DDD", "Telefone celular"], ["Endereço", "Endereço completo com CEP"], ["Local de nascimento", "Informações gerais"],
    ["Onde e com quem mora", "Em que cidade mora?"], ["Escolaridade", "Qual sua escolaridade?"], ["Ocupação", "Qual a sua ocupação?"],
    ["Religião", "Você pratica alguma religião?"], ["Peso e altura", "Saúde atual"]]],
  ["O que te traz à consulta", [
    ["Motivo", "Por favor descreva o principal motivo"], ["Há quanto tempo e o que já tentou", "Há quanto tempo isso acontece?"],
    ["Se nada mudar em 6 meses", "Se nada mudar nos próximos 6 meses"]]],
  ["Como isso afeta sua vida", [["Áreas", "Em quais áreas da sua vida"], ["Quanto atrapalha (0–10)", "De 0 a 10, quanto isso tem atrapalhado"]]],
  ["Nas últimas 2 semanas", [
    ["Pouco interesse ou prazer em fazer as coisas", "Durante as DUAS (2) SEMANAS"],
    ["Sentiu-se desanimado(a), deprimido(a) ou sem esperança", "Sentiu-se desanimado"],
    ["Sentiu-se mais irritado(a), mal-humorado(a) ou zangado(a) do que o usual", "Sentiu-se mais irritado"],
    ["Dormiu menos do que o usual, mas ainda com muita energia", "Dormiu menos do que o usual"],
    ["Começou muito mais projetos do que o usual ou fez coisas mais arriscadas", "Iniciou muito mais projetos"],
    ["Sentiu-se nervoso(a), ansioso(a), assustado(a), preocupado(a) ou tenso(a)", "Sentiu-se nervoso"],
    ["Sentiu pânico ou se sentiu amedrontado(a)", "Sentiu pânico"],
    ["Evitou situações que te deixam ansioso(a)", "Evitou situações"],
    ["Dores sem explicação (cabeça, costas, articulações, barriga, pernas)", "Dores e sofrimento sem explicação"],
    ["Sentiu que suas doenças não estão sendo levadas a sério", "Sentimento de que suas doenças"],
    ["Pensamentos de ferir a si mesmo(a)", "Pensamentos de ferir a si mesmo"],
    ["Ouviu coisas que outras pessoas não ouviam, como vozes", "Ouviu coisas que outras pessoas"],
    ["Sentiu que alguém podia ouvir seus pensamentos, ou que você ouvia os pensamentos de outra pessoa", "Sentiu que alguém podia ouvir"],
    ["Problemas com o sono", "Problemas com o sono"],
    ["Problemas de memória ou para se localizar (ex.: achar o caminho de casa)", "Problemas com a memória"],
    ["Pensamentos, impulsos ou imagens desagradáveis que voltam repetidamente", "Pensamentos, impulsos ou imagens"],
    ["Sentiu-se obrigado(a) a repetir certos comportamentos ou pensamentos", "Sentiu-se compelido"],
    ["Sentiu-se desligado(a) de si, do corpo, do ambiente ou das lembranças", "Sentiu-se desligado"],
    ["Sentiu-se sem saber quem você é ou o que quer da vida", "Sentiu-se sem saber quem"],
    ["Não se sentiu próximo(a) das pessoas nem aproveitou as relações", "Não se sentiu próximo"]], "rastreio"],
  ["Sono e rotina", [
    ["Horário de deitar e acordar", "Que horas você DEITA"], ["Queixa de sono", "Tem alguma queixa em relação ao seu sono"],
    ["Lazer", "O que você faz para se divertir"], ["Atividade física", "Você pratica atividade física"],
    ["Alimentos habituais", "Marque os alimentos"], ["Desconforto após comer", "Você tem algum desconforto após"],
    ["O que está certo na alimentação", "O que você acha que está CERTO"], ["O que está errado na alimentação", "O que você acha que está ERRADO"],
    ["Intestino", "Com que frequência você evacua"], ["Bebidas com cafeína por dia", "Número de bebidas cafeinadas"]]],
  ["Álcool, cigarro e outras substâncias", [
    ["Fuma", "Fuma tabaco"], ["Doses de álcool por semana", "Número de drinques com álcool"], ["Outras drogas", "Por favor liste o NOME de outras drogas"]]],
  ["Remédios e alergias", [["Remédios em uso", "Lista todas as medicações"], ["Alergias", "Lista as medicações e alimentos que você é alérgico"]]],
  ["Tratamentos anteriores", [
    ["Profissionais de saúde mental", "Por favor descreva sua experiência prévia"], ["Remédios psiquiátricos que já usou", "Por favor liste as medicações psiquiátricas"],
    ["Diagnósticos recebidos", "Você já recebeu algum diagnóstico"], ["Tentativa de suicídio", "Você já teve alguma tentativa de suicídio"],
    ["Internações psiquiátricas", "Você já precisou ser internado"]]],
  ["Saúde do corpo", [
    ["Problemas de saúde", "Por favor, liste todos os problemas de saúde"], ["Cirurgias", "Por favor liste as cirurgias"],
    ["Hospitalizações", "Por favor liste qualquer hospitalização"], ["Convulsão ou perda de consciência", "Você já sofreu alguma convulsão"],
    ["Detalhe", "Se sim para a pergunta anterior"]]],
  ["Infância e família", [
    ["Minha infância foi", "Minha infância foi"], ["Eventos estressantes", "Por favor descreva EVENTOS estressantes"],
    ["Pai", "Como você descreveria seu pai"], ["Mãe", "Como você descreveria sua mãe"], ["Jeito de ser", "Me descreva como você e as pessoas"]]],
  ["Trabalho e relações", [
    ["Trabalho: funções e desde quando", "Para me ajudar a entender melhor sua ocupação"], ["Estresse no trabalho", "O estresse no meu trabalho"],
    ["Felicidade no relacionamento", "Se você está em um relacionamento"], ["Tem amigos e apoio", "\"Eu tenho amigos suficientes"]]],
  ["Para mulheres", [
    ["Grávida", "Você está grávida"], ["Amamentando", "Você está amamentando"], ["Última menstruação", "Data da sua Última menstruação"],
    ["Gestações", "Número de gestações"], ["Filhos vivos", "Número de filhos vivos"], ["Humor com anticoncepcional ou antes da menstruação", "Por favor descreva qualquer PROBLEMA EMOCIONAL"],
    ["Método contraceptivo", "Você usa algum método contraceptivo"]]],
  ["Saúde mental na família", [["Problemas na família de sangue", "Por favor marque as alternativas"], ["Quem e o quê", "Se marcou qualquer item"]]],
  ["O que você espera", [
    ["De um bom atendimento", "O que você espera de um bom atendimento"], ["Se der certo, daqui a 3 meses", "Se o tratamento der certo"],
    ["O que poderia dificultar", "O que poderia dificultar seu acompanhamento"]]],
  ["Motivação", [["Importância de resolver agora (0–10)", "De 0 a 10, quão importante"], ["Confiança de mudar com ajuda (0–10)", "E de 0 a 10, quão confiante"]]]];

// Colunas que não viram texto: carimbo, e-mail, pontuação do Forms, colunas sem título e a de status.
const usadas = new Set([0]);
SECOES.forEach(([, itens]) => itens.forEach(([, inicio]) => colunas(inicio).forEach((i) => usadas.add(i))));
cab.forEach((h, i) => { if (!h || /^(endereco de e-mail|pontuacao|status v saude)$/.test(h)) usadas.add(i); });

const RISCO_TEXTO = /(suicid|me matar|se matar|tirar (a )?(minha|propria|sua) vida|nao quero mais viver|quer morrer|quero morrer|acabar com tudo|me ferir|me cortar|se cortar|se machuca)/;
const SEM_TENTATIVA = /^(nao|n|nunca|nenhuma|jamais|prefiro falar na consulta)\.?$/;
const LIVRES = ["Motivo", "Há quanto tempo e o que já tentou", "Se nada mudar em 6 meses", "Eventos estressantes", "Jeito de ser"];

function converter(vals) {
  const blocos = [], texto = {};
  for (const [titulo, itens, tipo] of SECOES) {
    const linhas = itens.map(([rot, inicio]) => {
      let v = valor(vals, inicio); texto[rot] = v;
      if (tipo === "rastreio") v = curtoFreq(v);
      return v || tipo === "rastreio" || titulo === "Identificação" || titulo === "Motivação" || titulo === "Como isso afeta sua vida" ? `${rot}: ${v || "—"}` : "";
    }).filter(Boolean);
    if (linhas.length || tipo === "rastreio") blocos.push([titulo, linhas.join("\n")]);
  }
  const outras = cab.map((h, i) => (usadas.has(i) || !String(vals[i] || "").trim() ? "" : `${dados.cabecalho[i].replace(/\s+/g, " ").trim()} ${String(vals[i]).trim()}`)).filter(Boolean);
  if (outras.length) blocos.push(["Outras respostas", outras.join("\n")]);

  const ferir = norm(curtoFreq(texto["Pensamentos de ferir a si mesmo(a)"]));
  const tentativa = norm(texto["Tentativa de suicídio"]);
  const risco = (ferir && ferir !== "nada" && ferir !== "—") || (tentativa && !SEM_TENTATIVA.test(tentativa)) ||
    RISCO_TEXTO.test(norm(LIVRES.map((k) => texto[k] || "").join(" ")));

  const cpf = String(texto["CPF"] || "").replace(/\D/g, "");
  const cabecalho = `[Adulto] ${texto["Nome completo"] || ""} · CPF ${cpf || "—"} · nasc. ${texto["Data de nascimento"] || "—"} · ${texto["Celular com DDD"] || ""}`;
  const corpo = (risco ? "⚠ ATENÇÃO: sinal de risco (autolesão/suicídio) nas respostas.\n\n" : "") +
    `PRÉ-CONSULTA · ADULTO (formulário oficial)\n${vals[0]}\n\n` + blocos.map(([t, x]) => `## ${t}\n${x || "—"}`).join("\n\n");
  return { cabecalho, corpo, risco };
}

// Cada resposta fica no MESMO número de linha da planilha de adulto (as outras linhas ficam vazias e o processar.js
// as ignora). Assim as chaves de "status" do saida.json já são as linhas onde gravar, e o e-mail de risco cita a linha certa.
const linhas = [["Carimbo de data/hora", "Paciente", "Respostas", "Status V Saúde"]], resumo = [];
for (const l of dados.linhas) {
  const { cabecalho, corpo, risco } = converter(l.valores);
  while (linhas.length < l.linha - 1) linhas.push(["", "", "", ""]);
  linhas[l.linha - 1] = [l.valores[0], cabecalho, corpo, (dados.status || {})[l.linha] || ""];
  resumo.push(`linha ${l.linha}: ${cabecalho}${risco ? " · ⚠ risco" : ""}`);
}
fs.writeFileSync(SAIDA, JSON.stringify({ inicio: dados.inicio, planilhaId: dados.planilhaId, linhas, links: dados.links || [] }, null, 2));
console.log(resumo.join("\n") || "nenhuma linha");
