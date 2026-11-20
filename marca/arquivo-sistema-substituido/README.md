# Marca Dr. Luis Felipe Jardim

> **SUBSTITUÍDO.** O sistema canônico é **"Dr. Luis Felipe Jardim — Marca"**: https://claude.ai/artifact/HkhxFoizd8u9aKWFqPi4Rn (valores Pantone/HEX oficiais do guia da Simple But, tema Oliva para telas do paciente, âmbar #E89925). Use aquele. Este fica só como registro e pode ser apagado; o âmbar #EC8C00 abaixo foi amostrado do vetor e difere do oficial.

Sistema visual digital do consultório: apps, páginas para pacientes, sites, e-books, slides de palestra e capas. Traduz para a tela o **Guia de Marca Pessoal** feito pela SIMPLE, but. (pasta "DR. LUIS FELIPE JARDIM + SIMPLE" no Drive), o arquivo vetorial `Símbolos.ai` e o ONE PAGE. Quando este sistema e o guia original divergirem, vale o guia original.

**Leia antes de criar qualquer material.** Todo app, slide ou e-book novo parte destes tokens. Não invente cores, não troque a ampulheta por outra forma e não use azul.

## Essência

Do ONE PAGE da Simple:

- **Atributos:** convicção, sabedoria, sensibilidade, inovação, atenção.
- **Proposta de valor:** responsabilidade (ética, rigor técnico-científico), inovação (conhecimento atualizado), cuidado e acolhimento (respeito ao indivíduo de forma única).
- **Propósito:** um diálogo sobre a jornada de cuidar da mente, para compreender a complexidade do viver, buscando sentido, equilíbrio e saúde emocional.
- **Manifesto (síntese):** "Sempre em construção. É preciso buscar simplicidade no cenário complexo." A ampulheta é o tempo que temos e o equilíbrio entre as duas metades.

## Tom de voz

- Fale como um médico que explica com calma, não como um folheto. Frases curtas, voz ativa, segunda pessoa ("você").
- Acolhedor sem ser infantil. Técnico quando o público é técnico; para pacientes, palavras do dia a dia ("crise de medo", não "ataque de pânico com agorafobia").
- Nunca alarmista. Em temas de risco, diga o que fazer e para onde ligar (CVV 188, SAMU 192), sempre.
- Assinatura escrita: **Dr. Luis Felipe Jardim** (sem acento em "Luis", como no logotipo). Médico psiquiatra e psicoterapeuta.
- Registros: CREMESC 12767 | RQE 22739 (psiquiatria) · RQE 24818 (psicoterapia). No RS: CREMERS 27789 | RQE 40903 · 43409.

## Cores

Nove cores oficiais + preto, amostradas do `Símbolos.ai`. Cada uma tem papel definido; o modo escuro tem valores próprios em `tokens.json`.

| Cor | Hex | Papel |
| --- | --- | --- |
| Oliva | `#656635` | **Cor-mãe.** Faixas de cabeçalho, botões principais, títulos. Fundo da capa do app. |
| Âmbar | `#EC8C00` | **O símbolo.** Ampulheta sobre oliva, detalhes, um destaque por tela. Nunca texto pequeno sobre fundo claro (use `ambar-texto`). |
| Pinho | `#38584B` | Estado "bem/leve" no semáforo clínico; ícones; segunda cor de blocos. |
| Terracota | `#C3622D` | Fundo de capas de impacto (ONE PAGE). No semáforo, o forte usa `alerta` (terracota fechado). |
| Castanho | `#83462D` | Cor do logotipo sobre fundo claro. Títulos editoriais em e-books. |
| Bronze | `#88764F` | Ícones secundários, rótulos neutros. |
| Areia | `#AA9071` | Fundo da tela de abertura do app; blocos de respiro. |
| Sálvia | `#A3B2A4` | Fundos suaves, gráficos de apoio. Não usar com texto branco. |
| Off-white | `#F2F1F0` | Texto sobre oliva/terracota; base dos fundos claros. |

**Combinações aprovadas:** off-white sobre oliva (5,3:1), âmbar sobre oliva (símbolo, nunca texto), castanho sobre off-white, oliva sobre off-white (5,4:1), off-white sobre terracota em títulos grandes.

**Semáforo clínico** (escalas, alertas, evolução): `ok` pinho → `moderado` âmbar fechado → `forte` terracota fechado. As três diferem em claridade, não só em tom, para quem tem daltonismo. Não use vermelho e verde puros.

## Tipografia

- **Títulos: Optima** (fonte oficial, OptimaLTPro). Já vem no iPhone e no Mac. Em Android/Windows a página cai para **Marcellus** (Google Fonts), a mais próxima no traço. Peso 600 nos títulos de tela, 400 em títulos editoriais grandes.
- **Texto: Garet** é a oficial; nas telas usamos **Figtree** (Google Fonts), geométrica e muito legível em celular. Pesos 400/600/700/800.
- Rótulos em caixa alta com espaçamento 0,12–0,14em (como "MÉDICO PSIQUIATRA" no logotipo).
- Escala de tela: 28 título · 20 subtítulo · 16 item · 15 texto · 13 nota · 11 rótulo.

Stack CSS:

```css
--font-display: Optima, "Optima nova", "Marcellus", Candara, serif;
--font-sans: "Figtree", system-ui, -apple-system, "Segoe UI", sans-serif;
```

## Símbolo e logotipo

- **Ampulheta facetada** (arquivo oficial, grupo *Símbolo*). Âmbar sobre oliva é a assinatura; oliva sobre areia é a variação da tela de abertura. Altura mínima em tela: 20px.
- A ampulheta **arredondada** (curvas suaves, cor salmão) usada em alguns slides antigos **não é a oficial**. Substituir pela facetada.
- Os **símbolos derivados** do `Símbolos.ai` (meia-ampulheta, cápsula, ampulheta inclinada) podem compor faixas decorativas, como no rodapé do ONE PAGE, sempre nas cores da paleta.
- **Logotipo**: nome em Optima Bold + "MÉDICO PSIQUIATRA" + registros. Castanho sobre claro, branco sobre oliva/terracota. Não redesenhar, não esticar, não aplicar sombra.

## Padrões de interface

- **Faixa de marca**: todo app abre com uma faixa oliva (símbolo âmbar + nome em Optima off-white + "PSIQUIATRIA · PSICOTERAPIA" em rótulo). É o elo visual com o app da App Store.
- **Fundo** off-white quente (`fundo`), cartões `superficie` com borda `linha` de 1px e raio 16px. Sombra quase imperceptível.
- **Botão principal** oliva com texto branco, raio 14px, largura total no celular.
- **Seleção** (chips, opções): borda `linha`; selecionado vira oliva cheio.
- **Alertas de risco**: fundo `alerta-fundo`, texto `alerta-texto`, com 188/192 em negrito.
- **Modo escuro**: fundo oliva quase preto (`#1B1C14`), não cinza-azulado.
- Marcadores de seção: uma mini-ampulheta âmbar antes do rótulo em caixa alta.

## Slides e e-books

- Capa: fundo oliva, ampulheta âmbar, nome em Optima off-white (como a "imagem fundo" da Simple). Alternativa de impacto: fundo terracota (ONE PAGE).
- Miolo: fundo off-white, títulos em Optima oliva ou castanho, texto Figtree/Garet.
- Destaques numéricos em âmbar sobre oliva; gráficos com pinho, âmbar, terracota, sálvia, areia (nessa ordem).
- QR codes: módulos `tinta` sobre branco, dentro de um cartão; legenda em Optima.
- Logos de parceiros (Sicoob etc.) ficam pequenos, no rodapé, nunca disputando com a ampulheta.

## Implementação no código

O repositório `regua-sintomas` tem o arquivo pronto `marca/marca.css` com todos os tokens (claro e escuro) e `marca/simbolo.svg`. Em artifacts, copiar o bloco `:root` daquele arquivo para dentro do `<style>`.

## Inventário dos materiais (status)

| Material | Status |
| --- | --- |
| Régua de sintomas e escalas (GitHub Pages) | Aplicado |
| Meu peso (GitHub Pages) | Aplicado |
| Central de Apps + QR codes | Aplicado |
| Receitas, Régua (artifact), Monitor de Peso, Escetamina, Cetamina, Psicoeducação, Contratos, Contrato (paciente), Secretária, Faturamento, Site Neuromodulação, Meu Treino | A aplicar |
| Palestras (slides), e-books, capas de rádio | A aplicar |
