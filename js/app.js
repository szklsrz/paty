'use strict';

/* ==========================================================================
   Dados de referência
   ========================================================================== */

// Fonte: index.xlsx, abas "Referências idade" (n = 427) e "Referências IMC" (n = 428). Linha: [n, P25, P50, P75].
const AGE_REF = {
  M: [[48, 4775.9, 7074.8, 8466.1], [37, 3303.3, 4635.7, 6473.3], [31, 4858.3, 5713.6, 7399], [31, 2418.1, 4684.7, 6840], [32, 3619.4, 5394.9, 7701], [31, 2807.3, 3937.3, 5085.8]],
  F: [[56, 4231.6, 5356.6, 6977.7], [35, 2676.6, 3537.6, 6124.6], [32, 3488, 3834.4, 5833.9], [31, 2981.9, 4281.1, 6459.9], [32, 2214.8, 3470.4, 5082.9], [31, 2000, 2721.9, 4782.2]],
};

const BMI_REF = {
  M: [[0, null, null, null], [79, 3921.9, 5877.1, 8583.5], [93, 3320.1, 5432.7, 7403.3], [28, 2782.8, 4035.7, 5169.1], [5, 3930, 5412.1, 6022], [5, 2603.9, 2996.4, 5001.7]],
  F: [[8, 2364.9, 4389.6, 5747], [96, 3004.1, 4615, 7115.7], [57, 2790.1, 3834.7, 6153.3], [35, 2748.4, 3523.3, 5599.5], [12, 2592.1, 3930.6, 6388.4], [10, 3082.8, 4605.1, 5554.5]],
};

// Médias da Tabela 2 do artigo (n = 433); usadas apenas no gráfico.
const AGE_MEANS = {
  M: [6396.7, 5139.7, 6588.2, 5160.2, 5788.4, 4446.2],
  F: [5470.4, 4571.4, 4680.5, 5106.8, 4091.0, 3923.3],
};

const AGE_KEYS = ['18–29', '30–39', '40–49', '50–59', '60–69', '≥70'];

/* ==========================================================================
   Textos dinâmicos (o texto fixo da página fica no HTML, em blocos lang)
   ========================================================================== */

const I18N = {
  pt: {
    locale: 'pt-BR',
    htmlLang: 'pt-BR',
    meta: {
      title: 'Calculadora de Passos Diários — compare sua média de passos',
      description: 'Calculadora educativa e gratuita: compare a sua média diária de passos com a de adultos brasileiros do mesmo sexo e idade que participaram de um estudo científico.',
    },
    placeholders: { age: 'Ex.: 45', weight: 'Ex.: 70,5', height: 'Ex.: 168', steps: 'Ex.: 6.500' },
    alts: {
      why1: 'Celular e relógio inteligente mostrando dados de atividade física',
      why2: 'Casal de idosos caminhando juntos na calçada',
      why3: 'Duas pessoas caminhando por uma trilha ensolarada',
      why4: 'Grupo de pessoas caminhando em uma rua arborizada',
      study: 'Pessoas caminhando em uma calçada arborizada da cidade',
      tips: 'Pessoa de tênis subindo uma escada de concreto',
    },
    aria: { mainNav: 'Navegação principal', mobileNav: 'Menu de navegação', footerNav: 'Navegação do rodapé' },
    sex: { M: 'Masculino', F: 'Feminino' },
    sexPlural: { M: 'homens', F: 'mulheres' },
    ageGroups: ['18 a 29 anos', '30 a 39 anos', '40 a 49 anos', '50 a 59 anos', '60 a 69 anos', '70 anos ou mais'],
    agePhrases: ['de 18 a 29 anos', 'de 30 a 39 anos', 'de 40 a 49 anos', 'de 50 a 59 anos', 'de 60 a 69 anos', 'com 70 anos ou mais'],
    bmiCategories: ['Abaixo do peso', 'Peso normal', 'Sobrepeso', 'Obesidade classe I', 'Obesidade classe II', 'Obesidade classe III'],
    bmiPhrases: ['abaixo do peso', 'com peso normal', 'com sobrepeso', 'com obesidade classe I', 'com obesidade classe II', 'com obesidade classe III'],
    levels: [
      null,
      { label: 'Faixa de baixo', text: 'Você deu menos passos que a maioria do seu grupo: pelo menos 3 em cada 4 pessoas caminharam mais.' },
      { label: 'Um pouco abaixo do meio', text: 'Você está um pouco abaixo do meio do seu grupo: mais da metade das pessoas caminhou mais que você.' },
      { label: 'Um pouco acima do meio', text: 'Você está acima do meio do seu grupo: caminhou mais que pelo menos metade das pessoas.' },
      { label: 'Entre os que mais caminham', text: 'Você está entre os que mais caminham no seu grupo: caminhou tanto quanto ou mais que 3 em cada 4 pessoas.' },
    ],
    bands: ['Não disponível', 'Abaixo do P25', 'Entre P25 e P50', 'Entre P50 e P75', 'Igual ou acima do P75'],
    sentences: [
      null,
      (g) => `Sua média está abaixo do P25 observado no grupo ${g}. Ela está na faixa inferior da distribuição desta amostra.`,
      (g) => `Sua média está entre P25 e P50 no grupo ${g}. Ela ficou abaixo da mediana observada.`,
      (g) => `Sua média está entre P50 e P75 no grupo ${g}. Ela ficou acima da mediana observada.`,
      (g) => `Sua média é igual ou superior ao P75 no grupo ${g}. Ela está na faixa superior da distribuição desta amostra.`,
    ],
    precision: {
      none: 'Sem participantes neste grupo',
      low: 'Baixa: grupo com menos de 20 participantes',
      exploratory: 'Exploratória: interpretar com cautela',
    },
    precisionShort: { none: 'Sem dados', low: 'Precisão baixa: poucas pessoas', exploratory: 'Interprete com cautela' },
    result: {
      comparedWith: (g) => `Comparado a ${g}`,
      groupSize: (n) => {
        if (n === 0) return 'Nenhuma pessoa deste grupo no estudo';
        return n === 1 ? '1 pessoa deste grupo no estudo' : `${n} pessoas deste grupo no estudo`;
      },
      noData: 'Comparação indisponível',
      noDataText: 'Nenhuma pessoa deste grupo participou do estudo, então não é possível fazer esta comparação.',
      you: 'Você',
      barAria: (steps, label) => `Sua média de ${steps} passos por dia fica na faixa “${label}”.`,
      legend: [
        (p) => `1 em cada 4 pessoas deu menos de ${p} passos por dia.`,
        (p) => `Metade deu menos de ${p}.`,
        (p) => `3 em cada 4 deram menos de ${p}.`,
      ],
      live: (g, label) => `Resultado pronto. Comparado a ${g}: ${label}.`,
    },
    tech: {
      group: 'Grupo de comparação',
      groupBmi: 'Grupo exploratório',
      n: 'n do grupo',
      p25: 'P25',
      p50: 'P50 (mediana)',
      p75: 'P75',
      position: 'Posição na distribuição observada',
      band: 'Faixa percentilar',
      precision: 'Precisão',
      sentence: 'Interpretação',
    },
    errors: {
      sex: 'Escolha uma das opções.',
      consent: 'Marque a caixa para consentir com o envio dos dados.',
      ageEmpty: 'Digite sua idade em anos, só com números (ex.: 45).',
      ageMin: 'A calculadora é para adultos, a partir de 18 anos.',
      ageMax: 'Confira a idade: use um valor entre 18 e 120 anos.',
      weight: 'Digite seu peso em quilos (ex.: 70,5).',
      height: 'Digite sua altura entre 100 e 250 cm (ex.: 168).',
      steps: 'Digite sua média de passos por dia, de 0 a 100.000.',
    },
    formIncomplete: 'Preencha os campos obrigatórios e marque o consentimento para ver o resultado.',
    sendError: 'Não foi possível enviar os dados. O resultado foi calculado, mas não enviado.',
    sendUnconfirmed: 'Solicitação enviada; não é possível confirmar o recebimento.',
    helper: {
      empty: 'Digite os passos de pelo menos um dia.',
      avg: (n, v) => `Média de ${n} ${n === 1 ? 'dia' : 'dias'}: ${v} passos`,
    },
    chart: {
      labels: ['18–29', '30–39', '40–49', '50–59', '60–69', '70+'],
      men: 'Homens',
      women: 'Mulheres',
      x: 'Faixa etária (anos)',
      y: 'Média de passos por dia',
      unit: 'passos/dia',
      aria: (men, women) => `Gráfico de barras com a média de passos por dia por faixa etária. Homens: ${men}. Mulheres: ${women}.`,
    },
  },

  en: {
    locale: 'en-US',
    htmlLang: 'en',
    meta: {
      title: 'Daily Step Calculator — compare your average steps',
      description: 'Free educational calculator: compare your average daily steps with Brazilian adults of the same sex and age who took part in a scientific study.',
    },
    placeholders: { age: 'e.g. 45', weight: 'e.g. 70.5', height: 'e.g. 168', steps: 'e.g. 6,500' },
    alts: {
      why1: 'Smartphone and smartwatch showing physical activity data',
      why2: 'Older couple walking together on a sidewalk',
      why3: 'Two people walking along a sunny path',
      why4: 'Group of people walking down a tree-lined street',
      study: 'People walking on a tree-lined city sidewalk',
      tips: 'Person in sneakers walking up concrete stairs',
    },
    aria: { mainNav: 'Main navigation', mobileNav: 'Navigation menu', footerNav: 'Footer navigation' },
    sex: { M: 'Male', F: 'Female' },
    sexPlural: { M: 'men', F: 'women' },
    ageGroups: ['18 to 29 years', '30 to 39 years', '40 to 49 years', '50 to 59 years', '60 to 69 years', '70 or older'],
    agePhrases: ['aged 18 to 29', 'aged 30 to 39', 'aged 40 to 49', 'aged 50 to 59', 'aged 60 to 69', 'aged 70 or older'],
    bmiCategories: ['Underweight', 'Normal weight', 'Overweight', 'Obesity class I', 'Obesity class II', 'Obesity class III'],
    bmiPhrases: ['who are underweight', 'with normal weight', 'who are overweight', 'with class I obesity', 'with class II obesity', 'with class III obesity'],
    levels: [
      null,
      { label: 'Lower range', text: 'You took fewer steps than most of your group: at least 3 in 4 people walked more.' },
      { label: 'Just below the middle', text: 'You are a little below the middle of your group: more than half of the people walked more than you.' },
      { label: 'Just above the middle', text: 'You are above the middle of your group: you walked more than at least half of the people.' },
      { label: 'Among the top walkers', text: 'You are among those who walk the most in your group: you walked as much as or more than 3 in 4 people.' },
    ],
    bands: ['Not available', 'Below P25', 'Between P25 and P50', 'Between P50 and P75', 'At or above P75'],
    sentences: [
      null,
      (g) => `Your average is below the P25 observed in the group ${g}. It falls in the lower part of this sample's distribution.`,
      (g) => `Your average is between P25 and P50 in the group ${g}. It is below the observed median.`,
      (g) => `Your average is between P50 and P75 in the group ${g}. It is above the observed median.`,
      (g) => `Your average is at or above P75 in the group ${g}. It falls in the upper part of this sample's distribution.`,
    ],
    precision: {
      none: 'No participants in this group',
      low: 'Low: group with fewer than 20 participants',
      exploratory: 'Exploratory: interpret with caution',
    },
    precisionShort: { none: 'No data', low: 'Low precision: few people', exploratory: 'Interpret with caution' },
    result: {
      comparedWith: (g) => `Compared with ${g}`,
      groupSize: (n) => {
        if (n === 0) return 'No one from this group in the study';
        return n === 1 ? '1 person from this group in the study' : `${n} people from this group in the study`;
      },
      noData: 'Comparison not available',
      noDataText: 'No one from this group took part in the study, so this comparison is not possible.',
      you: 'You',
      barAria: (steps, label) => `Your average of ${steps} steps a day falls in the “${label}” range.`,
      legend: [
        (p) => `1 in 4 people took fewer than ${p} steps a day.`,
        (p) => `Half took fewer than ${p}.`,
        (p) => `3 in 4 took fewer than ${p}.`,
      ],
      live: (g, label) => `Result ready. Compared with ${g}: ${label}.`,
    },
    tech: {
      group: 'Comparison group',
      groupBmi: 'Exploratory group',
      n: 'Group n',
      p25: 'P25',
      p50: 'P50 (median)',
      p75: 'P75',
      position: 'Position in the observed distribution',
      band: 'Percentile range',
      precision: 'Precision',
      sentence: 'Interpretation',
    },
    errors: {
      sex: 'Please choose one of the options.',
      consent: 'Check the box to consent to sending your details to Google Forms.',
      ageEmpty: 'Enter your age in years, numbers only (e.g. 45).',
      ageMin: 'The calculator is for adults aged 18 or older.',
      ageMax: 'Check your age: use a value between 18 and 120.',
      weight: 'Enter your weight in kilograms (e.g. 70.5).',
      height: 'Enter your height between 100 and 250 cm (e.g. 168).',
      steps: 'Enter your average steps per day, from 0 to 100,000.',
    },
    formIncomplete: 'Fill in the required fields and check the consent box to see your result.',
    sendError: 'Could not send your details to Google Forms. The result was calculated but not sent.',
    sendUnconfirmed: 'Request sent to Google Forms; receipt cannot be confirmed.',
    helper: {
      empty: 'Enter the steps for at least one day.',
      avg: (n, v) => `Average of ${n} ${n === 1 ? 'day' : 'days'}: ${v} steps`,
    },
    chart: {
      labels: ['18–29', '30–39', '40–49', '50–59', '60–69', '70+'],
      men: 'Men',
      women: 'Women',
      x: 'Age group (years)',
      y: 'Average steps per day',
      unit: 'steps/day',
      aria: (men, women) => `Bar chart of average steps per day by age group. Men: ${men}. Women: ${women}.`,
    },
  },
};

/* ==========================================================================
   Leitura dos campos (aceita formatos pt-BR e en)
   ========================================================================== */

function parseInteger(raw) {
  const s = String(raw).trim();
  return /^\d+$/.test(s) ? Number(s) : NaN;
}

function parseDecimal(raw) {
  const s = String(raw).trim();
  return /^\d+([.,]\d+)?$/.test(s) ? Number(s.replace(',', '.')) : NaN;
}

// "8.500", "8,500" e "8 500" são separadores de milhar; "6234,5" é decimal e é arredondado.
function parseSteps(raw) {
  const s = String(raw).trim();
  if (/^\d+$/.test(s)) return Number(s);
  if (/^\d{1,3}([.,\s]\d{3})+$/.test(s)) return Number(s.replace(/[.,\s]/g, ''));
  if (/^\d+[.,]\d{1,2}$/.test(s)) return Math.round(Number(s.replace(',', '.')));
  return NaN;
}

/* ==========================================================================
   Cálculo — réplica das fórmulas da aba "Calculadora" do index.xlsx
   ========================================================================== */

function ageGroupIndex(age) {
  if (age <= 29) return 0;
  if (age <= 39) return 1;
  if (age <= 49) return 2;
  if (age <= 59) return 3;
  if (age <= 69) return 4;
  return 5;
}

function bmiGroupIndex(bmi) {
  if (bmi < 18.5) return 0;
  if (bmi < 25) return 1;
  if (bmi < 30) return 2;
  if (bmi < 35) return 3;
  if (bmi < 40) return 4;
  return 5;
}

// 0 = sem referência ("Não disponível"); 1–4 = abaixo de P25, P25–P50, P50–P75, P75 ou mais.
function percentileBand(steps, ref) {
  if (ref.p25 === null) return 0;
  if (steps < ref.p25) return 1;
  if (steps < ref.p50) return 2;
  if (steps < ref.p75) return 3;
  return 4;
}

function precisionKey(n) {
  if (n === 0) return 'none';
  if (n < 20) return 'low';
  return 'exploratory';
}

function toRef([n, p25, p50, p75]) {
  return { n, p25, p50, p75 };
}

function calculate({ sex, age, weight, height, steps }) {
  const bmi = weight / (height / 100) ** 2;
  const ageIdx = ageGroupIndex(age);
  const bmiIdx = bmiGroupIndex(bmi);
  const ageRef = toRef(AGE_REF[sex][ageIdx]);
  const bmiRef = toRef(BMI_REF[sex][bmiIdx]);
  return {
    sex,
    steps,
    bmi,
    ageIdx,
    bmiIdx,
    ageRef,
    bmiRef,
    ageBand: percentileBand(steps, ageRef),
    bmiBand: percentileBand(steps, bmiRef),
  };
}

/* ==========================================================================
   Estado e utilidades de interface
   ========================================================================== */

let lang = document.documentElement.dataset.lang === 'en' ? 'en' : 'pt';
let lastResult = null;
let currentErrors = {};
let helperValue = null;
let chart = null;

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));
const dict = () => I18N[lang];
const reducedMotion = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;
const scrollBehavior = () => (reducedMotion() ? 'auto' : 'smooth');

function fmt(value, digits = 0) {
  return new Intl.NumberFormat(dict().locale, { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(value);
}

const form = $('#calc-form');
const fields = { age: $('#age'), weight: $('#weight'), height: $('#height'), steps: $('#steps') };
const consentCheckbox = $('#consent');
const sexGroup = $('#sex-group');
const formStatus = $('#form-status');
const resultBox = $('#resultado');
const resultTitle = $('#result-title');
const resultLive = $('#result-live');
const cards = { age: $('#compare-age'), bmi: $('#compare-bmi') };
const dayInputs = $$('[data-day]');
const helperAvg = $('#helper-avg');
const helperNote = $('#helper-note');
const helperUse = $('#helper-use');
const header = $('#site-header');
const menuButton = $('#menu-toggle');
const mobileMenu = $('#mobile-menu');
const chartCanvas = $('#chart-age');
const metaDescription = $('meta[name="description"]');

/* ==========================================================================
   Formulário
   ========================================================================== */

function readForm() {
  const errors = {};
  const sex = form.elements.sex.value;
  if (sex !== 'M' && sex !== 'F') errors.sex = 'sex';
  if (!consentCheckbox.checked) errors.consent = 'consent';

  const age = parseInteger(fields.age.value);
  if (Number.isNaN(age)) errors.age = 'ageEmpty';
  else if (age < 18) errors.age = 'ageMin';
  else if (age > 120) errors.age = 'ageMax';

  const weight = parseDecimal(fields.weight.value);
  if (!(weight >= 20 && weight <= 400)) errors.weight = 'weight';

  let height = parseDecimal(fields.height.value);
  if (height > 0 && height < 3) height = Number((height * 100).toFixed(4));
  if (!(height >= 100 && height <= 250)) errors.height = 'height';

  const steps = parseSteps(fields.steps.value);
  if (!(steps >= 0 && steps <= 100000)) errors.steps = 'steps';

  return { values: { sex, age, weight, height, steps }, errors };
}

function showErrors(errors) {
  currentErrors = errors;
  const messages = dict().errors;
  for (const name of ['sex', 'age', 'weight', 'height', 'steps']) {
    const box = $(`#err-${name}`);
    const key = errors[name];
    box.hidden = !key;
    box.textContent = key ? messages[key] : '';
    if (name === 'sex') sexGroup.dataset.invalid = String(Boolean(key));
    else fields[name].setAttribute('aria-invalid', String(Boolean(key)));
  }
  const consentError = errors.consent;
  const consentErrorBox = $('#err-consent');
  consentErrorBox.hidden = !consentError;
  consentErrorBox.textContent = consentError ? messages[consentError] : '';
  consentCheckbox.setAttribute('aria-invalid', String(Boolean(consentError)));
  const hasErrors = Object.keys(errors).length > 0;
  formStatus.hidden = !hasErrors;
  formStatus.textContent = hasErrors ? dict().formIncomplete : '';
}

function clearFieldError(name) {
  if (!currentErrors[name]) return;
  const rest = { ...currentErrors };
  delete rest[name];
  showErrors(rest);
}

function onSubmit(event) {
  event.preventDefault();
  const { values, errors } = readForm();
  showErrors(errors);

  const firstInvalid = ['consent', 'sex', 'age', 'weight', 'height', 'steps'].find((name) => errors[name]);
  if (firstInvalid) {
    lastResult = null;
    resultBox.hidden = true;
    if (firstInvalid === 'consent') consentCheckbox.focus();
    else (firstInvalid === 'sex' ? form.querySelector('input[name="sex"]') : fields[firstInvalid]).focus();
    return;
  }

  lastResult = calculate(values);
  sendResults(values);
  renderResult(true);
  resultBox.hidden = false;
  resultBox.scrollIntoView({ behavior: scrollBehavior(), block: 'start' });
  resultTitle.focus({ preventScroll: true });
}

async function sendResults({ sex, age, weight, height, steps }) {
  try {
    await fetch('https://docs.google.com/forms/d/e/1FAIpQLSc2TtpBPyXP_j225_5WQweYkBXWeOJYxt9TkCTXfh8Jn5rYqg/formResponse', {
      method: 'POST',
      mode: 'no-cors',
      body: new URLSearchParams({
        'entry.2021882982': String(age),
        'entry.1369367563': String(weight),
        'entry.531128076': String(height),
        'entry.917155943': String(steps),
        'entry.888947931': sex === 'M' ? 'Masculino' : 'Feminino',
      }),
    });
    formStatus.hidden = false;
    formStatus.textContent = dict().sendUnconfirmed;
  } catch (error) {
    console.error('Falha ao enviar os dados.', error);
    formStatus.hidden = false;
    formStatus.textContent = dict().sendError;
  }
}

function onReset() {
  // O evento "reset" dispara antes de os campos serem limpos.
  setTimeout(() => {
    showErrors({});
    lastResult = null;
    resultBox.hidden = true;
    updateHelper();
  }, 0);
}

/* ---------- Ajuda: média a partir dos 7 dias ---------- */

function updateHelper() {
  const values = [];
  for (const input of dayInputs) {
    const raw = input.value.trim();
    const value = parseSteps(raw);
    const valid = raw === '' || (value >= 0 && value <= 100000);
    input.setAttribute('aria-invalid', String(!valid));
    if (raw !== '' && valid) values.push(value);
  }
  helperValue = values.length ? Math.round(values.reduce((sum, v) => sum + v, 0) / values.length) : null;
  helperAvg.textContent = helperValue === null ? dict().helper.empty : dict().helper.avg(values.length, fmt(helperValue));
  helperNote.hidden = !(values.length > 0 && values.length < 7);
  helperUse.disabled = helperValue === null;
}

function useHelperAverage() {
  if (helperValue === null) return;
  fields.steps.value = fmt(helperValue);
  clearFieldError('steps');
  fields.steps.focus();
}

/* ==========================================================================
   Resultado
   ========================================================================== */

function groupPhrase(kind, r) {
  const d = dict();
  const phrase = kind === 'age' ? d.agePhrases[r.ageIdx] : d.bmiPhrases[r.bmiIdx];
  return `${d.sexPlural[r.sex]} ${phrase}`;
}

// Posição do marcador: cada faixa ocupa 1/4 da barra e o valor é interpolado dentro dela.
function markerPosition(steps, { p25, p50, p75 }, band) {
  const stops = [0, p25, p50, p75, p75 + (p75 - p25)];
  const low = stops[band - 1];
  const high = stops[band];
  const fraction = high > low ? Math.min(1, Math.max(0, (steps - low) / (high - low))) : 0.5;
  return Math.min(97, Math.max(3, (band - 1) * 25 + fraction * 25));
}

function fillDistribution(slot, ref, steps, band) {
  const d = dict();
  slot('marker').style.left = `${markerPosition(steps, ref, band)}%`;
  slot('markerLabel').textContent = d.result.you;
  slot('t25').textContent = fmt(ref.p25);
  slot('t50').textContent = fmt(ref.p50);
  slot('t75').textContent = fmt(ref.p75);
  slot('track').setAttribute('aria-label', d.result.barAria(fmt(steps), d.levels[band].label));
  slot('legend').replaceChildren(
    ...[ref.p25, ref.p50, ref.p75].map((value, i) => {
      const item = document.createElement('li');
      item.textContent = d.result.legend[i](fmt(value));
      return item;
    }),
  );
}

function fillTechnical(list, kind, r, ref, band) {
  const d = dict();
  const isAge = kind === 'age';
  const group = `${d.sex[r.sex]}, ${isAge ? AGE_KEYS[r.ageIdx] : d.bmiCategories[r.bmiIdx]}`;
  const value = (v) => (v === null ? '—' : fmt(v, 1));
  const rows = [
    [isAge ? d.tech.group : d.tech.groupBmi, group],
    [d.tech.n, String(ref.n)],
    [d.tech.p25, value(ref.p25)],
    [d.tech.p50, value(ref.p50)],
    [d.tech.p75, value(ref.p75)],
    [isAge ? d.tech.position : d.tech.band, d.bands[band]],
  ];
  if (isAge && band > 0) rows.push([d.tech.sentence, d.sentences[band](group)]);
  if (!isAge) rows.push([d.tech.precision, d.precision[precisionKey(ref.n)]]);

  list.replaceChildren(
    ...rows.flatMap(([term, description]) => {
      const dt = document.createElement('dt');
      const dd = document.createElement('dd');
      dt.textContent = term;
      dd.textContent = description;
      return [dt, dd];
    }),
  );
}

function fillComparison(kind, r) {
  const d = dict();
  const card = cards[kind];
  const slot = (name) => card.querySelector(`[data-slot="${name}"]`);
  const ref = kind === 'age' ? r.ageRef : r.bmiRef;
  const band = kind === 'age' ? r.ageBand : r.bmiBand;
  const hasData = band > 0;

  card.dataset.level = String(band);
  slot('title').textContent = d.result.comparedWith(groupPhrase(kind, r));
  slot('n').textContent = d.result.groupSize(ref.n);
  slot('badge').textContent = hasData ? d.levels[band].label : d.result.noData;
  slot('text').textContent = hasData ? d.levels[band].text : d.result.noDataText;
  slot('dist').hidden = !hasData;
  if (hasData) fillDistribution(slot, ref, r.steps, band);

  if (kind === 'bmi') {
    const precision = precisionKey(ref.n);
    slot('precision').dataset.precision = precision;
    slot('precision').textContent = d.precisionShort[precision];
  }
  fillTechnical(slot('tech'), kind, r, ref, band);
}

function renderResult(announce) {
  const r = lastResult;
  if (!r) return;
  const d = dict();
  const slot = (name) => resultBox.querySelector(`[data-slot="${name}"]`);
  slot('steps').textContent = fmt(r.steps);
  slot('ageGroup').textContent = d.ageGroups[r.ageIdx];
  slot('bmi').textContent = fmt(r.bmi, 1);
  slot('bmiCategory').textContent = d.bmiCategories[r.bmiIdx];
  fillComparison('age', r);
  fillComparison('bmi', r);
  if (announce) resultLive.textContent = d.result.live(groupPhrase('age', r), d.levels[r.ageBand].label);
}

/* ==========================================================================
   Gráfico (Chart.js), criado quando a seção aparece na tela
   ========================================================================== */

function chartColors() {
  const css = getComputedStyle(document.documentElement);
  const read = (name, fallback) => css.getPropertyValue(name).trim() || fallback;
  return { men: read('--color-brand-600', '#2b7045'), women: read('--color-warm-400', '#eea94f') };
}

function renderChart() {
  if (!chartCanvas || !window.Chart) return;
  const d = dict();
  const colors = chartColors();
  const summary = (values) => d.chart.labels.map((label, i) => `${label}: ${fmt(values[i])}`).join('; ');
  chartCanvas.setAttribute('aria-label', d.chart.aria(summary(AGE_MEANS.M), summary(AGE_MEANS.F)));

  const datasets = [
    { label: d.chart.men, data: AGE_MEANS.M, backgroundColor: colors.men, borderRadius: 8, maxBarThickness: 40 },
    { label: d.chart.women, data: AGE_MEANS.F, backgroundColor: colors.women, borderRadius: 8, maxBarThickness: 40 },
  ];

  if (chart) {
    chart.data.labels = d.chart.labels;
    chart.data.datasets.forEach((dataset, i) => { dataset.label = datasets[i].label; });
    chart.options.scales.x.title.text = d.chart.x;
    chart.options.scales.y.title.text = d.chart.y;
    chart.update();
    return;
  }

  window.Chart.defaults.font.family = 'Inter, ui-sans-serif, system-ui, sans-serif';
  window.Chart.defaults.color = '#475569';
  chart = new window.Chart(chartCanvas, {
    type: 'bar',
    data: { labels: d.chart.labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: reducedMotion() ? false : { duration: 900 },
      plugins: {
        legend: { display: false },
        tooltip: {
          padding: 12,
          callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)} ${dict().chart.unit}` },
        },
      },
      scales: {
        x: { grid: { display: false }, title: { display: true, text: d.chart.x } },
        y: {
          beginAtZero: true,
          grid: { color: '#e2e8f0' },
          ticks: { callback: (value) => fmt(value) },
          title: { display: true, text: d.chart.y },
        },
      },
    },
  });
}

function initChart() {
  if (!chartCanvas) return;
  if (!('IntersectionObserver' in window)) {
    renderChart();
    return;
  }
  const observer = new IntersectionObserver((entries) => {
    if (entries.some((entry) => entry.isIntersecting)) {
      renderChart();
      observer.disconnect();
    }
  }, { rootMargin: '200px' });
  observer.observe(chartCanvas);
}

/* ==========================================================================
   Idioma
   ========================================================================== */

function applyLanguage(next, persist) {
  lang = next;
  const d = dict();
  const root = document.documentElement;
  root.dataset.lang = lang;
  root.lang = d.htmlLang;
  document.title = d.meta.title;
  metaDescription?.setAttribute('content', d.meta.description);
  $$('[data-i18n-placeholder]').forEach((el) => { el.placeholder = d.placeholders[el.dataset.i18nPlaceholder]; });
  $$('[data-i18n-alt]').forEach((el) => { el.alt = d.alts[el.dataset.i18nAlt]; });
  $$('[data-i18n-aria]').forEach((el) => { el.setAttribute('aria-label', d.aria[el.dataset.i18nAria]); });

  if (persist) {
    try {
      localStorage.setItem('lang', lang);
    } catch {
      // Armazenamento bloqueado: o idioma vale só nesta visita.
    }
  }

  showErrors(currentErrors);
  updateHelper();
  renderResult(false);
  if (chart) renderChart();
}

/* ==========================================================================
   Cabeçalho, menu e animações
   ========================================================================== */

function setMenu(open) {
  menuButton.setAttribute('aria-expanded', String(open));
  mobileMenu.hidden = !open;
  header.classList.toggle('is-open', open);
}

function onScroll() {
  header.classList.toggle('is-scrolled', window.scrollY > 24);
}

function initReveal() {
  const items = $$('.reveal');
  if (!('IntersectionObserver' in window) || reducedMotion()) {
    items.forEach((el) => el.classList.add('is-visible'));
    return;
  }
  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (!entry.isIntersecting) return;
      entry.target.classList.add('is-visible');
      observer.unobserve(entry.target);
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
  items.forEach((el) => observer.observe(el));
}

/* ==========================================================================
   Inicialização
   ========================================================================== */

form.addEventListener('submit', onSubmit);
form.addEventListener('reset', onReset);
Object.entries(fields).forEach(([name, input]) => input.addEventListener('input', () => clearFieldError(name)));
$$('input[name="sex"]').forEach((radio) => radio.addEventListener('change', () => clearFieldError('sex')));
consentCheckbox.addEventListener('change', () => clearFieldError('consent'));
dayInputs.forEach((input) => input.addEventListener('input', updateHelper));
helperUse.addEventListener('click', useHelperAverage);

$('#recalc').addEventListener('click', () => {
  form.scrollIntoView({ behavior: scrollBehavior(), block: 'start' });
  fields.age.focus({ preventScroll: true });
});

$$('[data-lang-toggle]').forEach((button) => {
  button.addEventListener('click', () => applyLanguage(lang === 'pt' ? 'en' : 'pt', true));
});

menuButton.addEventListener('click', () => setMenu(menuButton.getAttribute('aria-expanded') !== 'true'));
mobileMenu.addEventListener('click', (event) => {
  if (event.target.closest('a')) setMenu(false);
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && menuButton.getAttribute('aria-expanded') === 'true') {
    setMenu(false);
    menuButton.focus();
  }
});

window.addEventListener('scroll', onScroll, { passive: true });
$$('[data-year]').forEach((el) => { el.textContent = String(new Date().getFullYear()); });

applyLanguage(lang, false);
window.lucide?.createIcons();
initReveal();
initChart();
onScroll();
