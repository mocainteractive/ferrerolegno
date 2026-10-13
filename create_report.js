const pptxgen = require('pptxgenjs');
const pres = new pptxgen();

pres.layout = 'LAYOUT_WIDE';
pres.title = 'Ferrero Legno - Report Settembre 2026';
pres.subject = 'Report mensile settembre 2026';
pres.author = 'moca interactive';
pres.company = 'moca interactive';

pres.theme = { headFontFace: 'Figtree', bodyFontFace: 'Figtree' };

const THEME = {
  name: 'Ferrero Legno Report',
  headFontFace: 'Figtree',
  bodyFontFace: 'Figtree',
  colors: {
    dk1: '191919', lt1: 'FFFFFF', dk2: '595959', lt2: 'F4F5F6',
    accent1: 'E52217', accent2: '118541', accent3: '5781FF',
    accent4: '551FC4', accent5: 'CE2A59', accent6: 'DB5E29',
    hlink: 'E52217', folHlink: 'CE2A59'
  }
};

const RED = 'E52217';
const WHITE = 'FFFFFF';
const LGRAY = 'F4F5F6';
const DARK = '191919';
const GRAY = '595959';
const GREEN = '118541';
const BLUE = '5781FF';
const PURPLE = '551FC4';
const PINK = 'CE2A59';
const ORANGE = 'DB5E29';
const DGRAY = '484848';
const FONT = 'Figtree';

function addSlideBase(slide, label) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.04, fill: { color: RED } });
  if (label) {
    slide.addText(label, {
      x: 0.5, y: 0.35, w: 10, h: 0.3,
      fontSize: 10, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
    });
  }
  slide.addShape(pres.shapes.OVAL, { x: 12.35, y: 7.0, w: 0.18, h: 0.18, fill: { color: RED } });
  slide.addText('moca', {
    x: 12.57, y: 6.95, w: 0.55, h: 0.25,
    fontSize: 10, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 7.05, w: 9, h: 0.25,
    fontSize: 7, fontFace: FONT, color: 'AAAAAA', isTextBox: true, margin: 0
  });
}

function addTitle(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 0.7, w: 12, h: 0.75,
    fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
}

// =============================================
// SLIDE 1 - COVER
// =============================================
pres.addSection({ title: 'Cover' });
const s1 = pres.addSlide({ sectionTitle: 'Cover' });

s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 0.05, fill: { color: RED } });

s1.addText([
  { text: 'FERRERO', options: { fontSize: 56, fontFace: FONT, bold: true, color: WHITE, breakLine: true } },
  { text: 'LEGNO', options: { fontSize: 56, fontFace: FONT, bold: true, color: WHITE } }
], { x: 0.8, y: 1.0, w: 6, h: 2.6, isTextBox: true, margin: 0, lineSpacingMultiple: 0.85 });

s1.addText('DESIGN TO BE OPEN', {
  x: 0.8, y: 3.4, w: 5, h: 0.4,
  fontSize: 12, fontFace: FONT, color: '777777', charSpacing: 3, isTextBox: true, margin: 0
});

s1.addText('Report settembre 2026', {
  x: 0.8, y: 5.0, w: 8, h: 0.7,
  fontSize: 28, fontFace: FONT, color: WHITE, isTextBox: true, margin: 0
});

s1.addText('13 ottobre 2026', {
  x: 0.8, y: 5.7, w: 4, h: 0.4,
  fontSize: 13, fontFace: FONT, color: '777777', isTextBox: true, margin: 0
});

s1.addShape(pres.shapes.OVAL, { x: 0.8, y: 6.65, w: 0.22, h: 0.22, fill: { color: RED } });
s1.addText('moca', {
  x: 1.1, y: 6.6, w: 1.2, h: 0.3,
  fontSize: 14, fontFace: FONT, bold: true, color: WHITE, isTextBox: true, margin: 0
});

// =============================================
// SLIDE 2 - DIVIDER: ANDAMENTO E RISULTATI
// =============================================
pres.addSection({ title: 'Andamento e risultati' });
const s2 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s2.addText([
  { text: 'Andamento e risultati', options: { fontSize: 38, fontFace: FONT, color: WHITE, italic: true, breakLine: true } },
  { text: 'Settembre 2026.', options: { fontSize: 38, fontFace: FONT, color: '888888', italic: true } }
], { x: 1.0, y: 2.5, w: 11, h: 2.5, isTextBox: true, margin: 0 });

// =============================================
// SLIDE 3 - TRAFFICO | PANORAMICA (line chart + KPIs)
// =============================================
const s3 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s3, 'TRAFFICO AL SITO | PANORAMICA');
addTitle(s3, 'Forte crescita delle sessioni a settembre: +59% rispetto al 2025, trainata dai canali paid');

const days = Array.from({length: 30}, (_, i) => String(i + 1));
const sess2026 = [584,702,633,656,564,521,623,620,941,935,1046,938,1024,938,986,1076,985,952,969,955,966,953,975,937,885,892,910,907,894,863];
const sess2025 = [604,627,520,556,506,474,380,556,579,583,539,535,501,433,533,592,573,599,559,448,410,595,522,610,539,536,477,411,602,540];

s3.addChart(pres.charts.LINE, [
  { name: 'Settembre 2026', labels: days, values: sess2026 },
  { name: 'Settembre 2025', labels: days, values: sess2025 }
], {
  x: 0.3, y: 1.7, w: 8.5, h: 4.0,
  showTitle: false,
  showLegend: true, legendPos: 't', legendFontSize: 9, legendFontFace: '+mn-lt', legendColor: GRAY,
  chartColors: [RED, 'CCCCCC'],
  lineSize: 2.5, lineSmooth: true,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 7, catAxisLabelFontFace: '+mn-lt',
  valAxisLabelColor: GRAY, valAxisLabelFontSize: 8, valAxisLabelFontFace: '+mn-lt',
  valGridLine: { color: 'EEEEEE', size: 0.5 },
  catGridLine: { style: 'none' },
  showValue: false
});

const kpis = [
  { label: 'SESSIONI', val: '25.691', delta: '+59,4%', color: GREEN, sub: 'vs 16.115' },
  { label: 'UTENTI', val: '18.350', delta: '+69,7%', color: GREEN, sub: 'vs 10.814' },
  { label: 'VISUALIZZAZIONI', val: '84.993', delta: '+9,1%', color: GREEN, sub: 'vs 77.872' },
  { label: 'ENGAGEMENT', val: '86,0%', delta: '-0,1pp', color: GRAY, sub: 'vs 86,1%' }
];

kpis.forEach((kpi, i) => {
  const yBase = 1.7 + i * 1.05;
  const xBase = 9.2;
  s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: xBase, y: yBase, w: 3.8, h: 0.9, rectRadius: 0.08, fill: { color: LGRAY }
  });
  s3.addText(kpi.label, {
    x: xBase + 0.15, y: yBase + 0.08, w: 2, h: 0.2,
    fontSize: 7, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0
  });
  s3.addText(kpi.val, {
    x: xBase + 0.15, y: yBase + 0.3, w: 1.8, h: 0.4,
    fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s3.addText(kpi.delta, {
    x: xBase + 2.0, y: yBase + 0.3, w: 0.9, h: 0.4,
    fontSize: 13, fontFace: FONT, bold: true, color: kpi.color, isTextBox: true, margin: 0
  });
  s3.addText(kpi.sub, {
    x: xBase + 2.9, y: yBase + 0.4, w: 0.8, h: 0.2,
    fontSize: 7, fontFace: FONT, color: 'AAAAAA', isTextBox: true, margin: 0
  });
});

s3.addText('Il traffico cresce grazie ai canali paid, che portano il 70% delle sessioni totali. Le sessioni giornaliere raggiungono picchi sopra le 1.000 dalla seconda settimana.', {
  x: 0.5, y: 6.0, w: 8.5, h: 0.7,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});
addFooter(s3, 'Fonte dati: GA4 | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 4 - TRAFFICO | MIX CANALI (doughnut + bar)
// =============================================
const s4 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s4, 'TRAFFICO AL SITO | MIX CANALI');
addTitle(s4, 'Il paid guida il 70% del traffico totale con Paid Search e Cross-network in fortissima crescita');

s4.addChart(pres.charts.DOUGHNUT, [
  { name: 'Canali', labels: ['Paid Search', 'Paid Social', 'Cross-network', 'Organic Search', 'Direct', 'Altro'], values: [8129, 5933, 3917, 4271, 1914, 1527] }
], {
  x: 0.3, y: 1.7, w: 5.0, h: 4.2,
  showTitle: false,
  showLegend: true, legendPos: 'b', legendFontSize: 8, legendFontFace: '+mn-lt', legendColor: GRAY,
  chartColors: [RED, BLUE, PURPLE, GREEN, ORANGE, DGRAY],
  dataLabelPosition: 'outEnd',
  showValue: false,
  showPercent: true,
  dataLabelFontSize: 9,
  dataLabelFontFace: '+mn-lt',
  dataLabelColor: DARK
});

s4.addChart(pres.charts.BAR, [
  { name: 'Set 2026', labels: ['Paid Search', 'Paid Social', 'Cross-net.', 'Organic', 'Direct'], values: [8129, 5933, 3917, 4271, 1914] },
  { name: 'Set 2025', labels: ['Paid Search', 'Paid Social', 'Cross-net.', 'Organic', 'Direct'], values: [3876, 4676, 19, 4448, 2009] }
], {
  x: 5.8, y: 1.7, w: 7.0, h: 4.2,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [RED, 'CCCCCC'],
  showValue: false,
  showLegend: true, legendPos: 't', legendFontSize: 8, legendFontFace: '+mn-lt', legendColor: GRAY,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 9, catAxisLabelFontFace: '+mn-lt',
  valAxisLabelColor: GRAY, valAxisLabelFontSize: 7, valAxisLabelFontFace: '+mn-lt',
  valGridLine: { color: 'EEEEEE', size: 0.5 },
  catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin'
});

s4.addText([
  { text: 'Paid Search (+110%) ', options: { bold: true, color: RED } },
  { text: 'e ', options: {} },
  { text: 'Cross-network (+20.516%) ', options: { bold: true, color: RED } },
  { text: 'trainano la crescita. Cross-network raccoglie Demand Gen, PMax, YouTube e Display. L\'organic cala del 4%, il direct del 4,7%.', options: {} }
], {
  x: 0.5, y: 6.2, w: 12.3, h: 0.5,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});
addFooter(s4, 'Fonte dati: GA4 | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 5 - ORGANICO | I NUMERI CHIAVE (BIG %)
// =============================================
const s5 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s5, 'TRAFFICO ORGANICO | I NUMERI CHIAVE');
addTitle(s5, 'L\'organico continua il trend di calo YoY, trainato dalla contrazione della domanda e dall\'erosione AI');

const bigNumbers = [
  {
    pct: '-18%', label: 'CLIC ORGANICI (IT)',
    desc: 'Da 6.593 a 5.391 clic.\nIl calo continua il trend dei mesi precedenti, legato alla contrazione della domanda e all\'effetto AI Overview.',
    color: RED
  },
  {
    pct: '+1,7%', label: 'IMPRESSION',
    desc: 'Da 121.681 a 123.803 impression.\nLe impression crescono leggermente, segno che il sito mantiene visibilità ma perde clic.',
    color: GREEN
  },
  {
    pct: '+4', label: 'POSIZIONI GUADAGNATE',
    desc: 'Da 15,9 a 11,9 posizione media.\nIl posizionamento migliora sensibilmente, ma non basta a compensare il calo dei volumi di ricerca.',
    color: GREEN
  }
];

bigNumbers.forEach((item, i) => {
  const x = 0.5 + i * 4.2;
  s5.addText(item.pct, {
    x: x, y: 1.9, w: 3.8, h: 1.8,
    fontSize: 72, fontFace: FONT, bold: true, color: item.color, isTextBox: true, margin: 0
  });
  s5.addText(item.label, {
    x: x, y: 3.7, w: 3.8, h: 0.3,
    fontSize: 11, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s5.addText(item.desc, {
    x: x, y: 4.1, w: 3.8, h: 1.6,
    fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3
  });
  if (i < 2) {
    s5.addShape(pres.shapes.RECTANGLE, {
      x: x + 4.05, y: 2.0, w: 0.02, h: 3.5, fill: { color: 'DDDDDD' }
    });
  }
});

addFooter(s5, 'Fonte dati: GSC - IT | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 6 - PERCHÉ IL CALO ORGANICO? (3-column narrative)
// =============================================
const s6 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s6, 'TRAFFICO ORGANICO | ANALISI');
addTitle(s6, 'Perché stiamo subendo questo calo?');

const reasons = [
  {
    num: '01', title: 'AI OVERVIEW',
    text: 'Google mostra risposte generate dall\'AI direttamente nei risultati di ricerca, riducendo i clic verso i siti web. I contenuti informativi, come quelli del magazine, sono i più colpiti. Il CTR medio scende da 5,42% a 4,35% (-1,07pp).'
  },
  {
    num: '02', title: 'CALO RICERCHE BRAND',
    text: 'Le ricerche legate al brand "Ferrero Legno" continuano a diminuire, in linea con un trend già visibile nei mesi precedenti. Questo riduce il volume di clic organici anche sulle pagine meglio posizionate.'
  },
  {
    num: '03', title: 'CONTRAZIONE DOMANDA',
    text: 'Il settore porte e infissi di design registra una contrazione delle ricerche online a livello generale. Nonostante il miglioramento del posizionamento (+4 posizioni), il volume complessivo di ricerche è inferiore rispetto al 2025.'
  }
];

reasons.forEach((r, i) => {
  const x = 0.5 + i * 4.2;
  s6.addShape(pres.shapes.OVAL, { x: x, y: 1.8, w: 0.55, h: 0.55, fill: { color: RED } });
  s6.addText(r.num, {
    x: x, y: 1.8, w: 0.55, h: 0.55,
    fontSize: 16, fontFace: FONT, bold: true, color: WHITE, align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });
  s6.addText(r.title, {
    x: x + 0.7, y: 1.85, w: 3.2, h: 0.4,
    fontSize: 14, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s6.addShape(pres.shapes.RECTANGLE, { x: x, y: 2.5, w: 3.8, h: 0.02, fill: { color: RED } });
  s6.addText(r.text, {
    x: x, y: 2.7, w: 3.8, h: 3.5,
    fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3
  });
});

addFooter(s6, 'Fonte dati: GSC, Google Trends | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 6B - ORGANICO | RICERCHE BRAND E TOP PAGES
// =============================================
const s6b = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s6b, 'TRAFFICO ORGANICO | APPROFONDIMENTO GSC');
addTitle(s6b, 'Le ricerche brand calano fino al -53%, ma il catalogo prodotti cresce del 32%');

s6b.addText('QUERY BRAND: CONFRONTO YOY', {
  x: 0.5, y: 1.7, w: 6, h: 0.3,
  fontSize: 10, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const brandQHeader = [[
  { text: 'Query', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Clic 2026', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Clic 2025', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Var. %', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const brandQData = [
  ['ferrero legno', '254', '537', '-52,7%'],
  ['ferrero porte', '241', '367', '-34,3%'],
  ['ferrerolegno', '97', '167', '-41,9%'],
  ['ferrero porte interne', '89', '61', '+45,9%'],
  ['ferrero legno porte', '103', '26', '+296,2%'],
  ['ferrero legno listino prezzi', '39', '63', '-38,1%'],
  ['ferrero legno maniglie', '27', '31', '-12,9%']
];
const brandQRows = [
  ...brandQHeader,
  ...brandQData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: {
      fontSize: 9, fontFace: FONT,
      color: j === 3 && cell.startsWith('+') ? GREEN : j === 3 && cell.startsWith('-') ? RED : DARK,
      align: j > 0 ? 'right' : 'left',
      fill: { color: i % 2 === 0 ? WHITE : LGRAY }
    }
  })))
];

s6b.addTable(brandQRows, {
  x: 0.5, y: 2.05, w: 6.0,
  border: { type: 'none' },
  colW: [2.4, 0.9, 0.9, 0.9],
  rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
  autoPage: false
});

s6b.addText('TOP PAGINE: VARIAZIONE CLIC', {
  x: 6.9, y: 1.7, w: 6, h: 0.3,
  fontSize: 10, fontFace: FONT, bold: true, color: BLUE, isTextBox: true, margin: 0
});

s6b.addChart(pres.charts.BAR, [
  { name: 'Set 2026', labels: ['Homepage', 'Tutte le porte', 'Soluzioni speciali', 'Punti vendita', 'Porte filo muro', 'Maniglie'], values: [631, 556, 260, 230, 183, 150] },
  { name: 'Set 2025', labels: ['Homepage', 'Tutte le porte', 'Soluzioni speciali', 'Punti vendita', 'Porte filo muro', 'Maniglie'], values: [1237, 422, 397, 170, 214, 180] }
], {
  x: 6.9, y: 2.05, w: 5.9, h: 3.0,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [BLUE, 'CCCCCC'],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 7, dataLabelFontFace: '+mn-lt',
  showLegend: true, legendPos: 't', legendFontSize: 8, legendFontFace: '+mn-lt', legendColor: GRAY,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 8, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin'
});

s6b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.2, w: 12.3, h: 1.4, rectRadius: 0.08, fill: { color: LGRAY } });
s6b.addText([
  { text: 'GSC aggregato: ', options: { bold: true, color: DARK } },
  { text: '6.738 clic (-13,3%) | 220.993 impression (+12,1%) | CTR 3,05% (-0,89pp) | Pos. media 10,3 (vs 14,3: +4 posizioni)', options: { color: GRAY } }
], {
  x: 0.7, y: 5.3, w: 11.8, h: 0.35,
  fontSize: 10, fontFace: FONT, isTextBox: true, margin: 0
});
s6b.addText([
  { text: 'Il calo delle ricerche brand (-53% su "ferrero legno") ', options: { bold: true } },
  { text: 'e compensato dalla crescita del catalogo: "tutte le porte" segna +32%, trainata dal paid. "Ferrero legno porte" e "ferrero porte interne" crescono grazie al posizionamento organico migliorato. Le impression globali crescono (+12%) nonostante il calo dei clic: il sito e piu visibile, ma l\'AI Overview erode i clic.', options: {} }
], {
  x: 0.7, y: 5.7, w: 11.8, h: 0.8,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3
});
addFooter(s6b, 'Fonte dati: Google Search Console | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 6C - CORRELAZIONE BRAND SEARCH E TRAFFICO DIRETTO
// =============================================
const s6c = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s6c, 'TRAFFICO ORGANICO | CORRELAZIONE');
addTitle(s6c, 'Le ricerche brand crollano fino al -53%, il traffico diretto segue a -4,7%: segnale di erosione della notorieta');

s6c.addText('CALO RICERCHE BRAND (GSC)', {
  x: 0.5, y: 1.7, w: 6, h: 0.3,
  fontSize: 10, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const brandDecline = [
  { query: '"ferrero legno"', val: '-52,7%', from: '537', to: '254' },
  { query: '"ferrero porte"', val: '-34,3%', from: '367', to: '241' },
  { query: '"ferrerolegno"', val: '-41,9%', from: '167', to: '97' }
];
brandDecline.forEach((b, i) => {
  const y = 2.15 + i * 1.1;
  s6c.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: y, w: 5.9, h: 0.95, rectRadius: 0.08, fill: { color: LGRAY } });
  s6c.addText(b.query, { x: 0.7, y: y + 0.08, w: 3.5, h: 0.25, fontSize: 11, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s6c.addText(b.val, { x: 4.2, y: y + 0.05, w: 2, h: 0.35, fontSize: 22, fontFace: FONT, bold: true, color: RED, align: 'right', isTextBox: true, margin: 0 });
  s6c.addText('da ' + b.from + ' a ' + b.to + ' clic', { x: 0.7, y: y + 0.45, w: 5, h: 0.25, fontSize: 9, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
});

s6c.addText('IMPATTO SUL TRAFFICO', {
  x: 6.8, y: 1.7, w: 6, h: 0.3,
  fontSize: 10, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const trafficImpact = [
  { label: 'TRAFFICO DIRETTO', val: '-4,7%', sub: '1.914 vs 2.009 sessioni', desc: 'Chi conosce il brand e naviga direttamente cala, in linea con la minore notorieta.' },
  { label: 'ORGANIC SEARCH', val: '-4,0%', sub: '4.271 vs 4.448 sessioni', desc: 'Il calo e contenuto grazie al miglioramento del posizionamento (+4 pos.), che compensa la domanda inferiore.' },
  { label: 'IMPRESSION GSC', val: '+12,1%', sub: '220.993 vs 197.214', desc: 'Il sito e piu visibile, ma l\'AI Overview intercetta i clic prima che arrivino al sito.' }
];
trafficImpact.forEach((t, i) => {
  const y = 2.15 + i * 1.1;
  s6c.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: y, w: 5.9, h: 0.95, rectRadius: 0.08, fill: { color: LGRAY } });
  s6c.addText(t.label, { x: 7.0, y: y + 0.05, w: 3, h: 0.25, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s6c.addText(t.val, { x: 10.4, y: y + 0.02, w: 2, h: 0.35, fontSize: 22, fontFace: FONT, bold: true, color: t.val.startsWith('+') ? GREEN : RED, align: 'right', isTextBox: true, margin: 0 });
  s6c.addText(t.sub, { x: 7.0, y: y + 0.35, w: 5.5, h: 0.2, fontSize: 9, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
  s6c.addText(t.desc, { x: 7.0, y: y + 0.55, w: 5.5, h: 0.3, fontSize: 9, fontFace: FONT, color: GRAY, italic: true, isTextBox: true, margin: 0 });
});

s6c.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.6, w: 12.3, h: 1.1, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s6c.addText([
  { text: 'La correlazione e chiara: ', options: { bold: true, color: RED } },
  { text: 'meno ricerche brand (-40/53%) significano meno clic organici e meno visite dirette. Tuttavia, il miglioramento del posizionamento e la crescita delle impression (+12,1%) suggeriscono che il sito sta guadagnando visibilita su query non-brand, compensando parzialmente il calo. L\'investimento paid diventa sempre piu strategico per mantenere il flusso di traffico.', options: { color: GRAY } }
], {
  x: 0.7, y: 5.65, w: 11.8, h: 1.0,
  fontSize: 10, fontFace: FONT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3
});
addFooter(s6c, 'Fonte dati: GSC, GA4 | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 7 - FOCUS ADVERTISING | PANORAMICA (2-column paid only)
// =============================================
const s7 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s7, 'TRAFFICO AL SITO | FOCUS ADVERTISING');
addTitle(s7, 'Google quadruplica i clic (+465%), Meta mantiene i volumi con CPC stabile');

const paidPlatforms = [
  {
    name: 'GOOGLE ADS', color: RED,
    metrics: [
      { k: 'Clic', v: '44.231', d: '+465%' },
      { k: 'Impression', v: '465.037', d: '+552%' },
      { k: 'Investimento', v: '5.193 €', d: '+102%' },
      { k: 'CPC medio', v: '0,12 €', d: '-64%' },
      { k: 'Quota traffico paid', v: '57,6%', d: '' }
    ]
  },
  {
    name: 'META ADS', color: BLUE,
    metrics: [
      { k: 'Link clic', v: '12.970', d: '-6,7%' },
      { k: 'Reach', v: '663.579', d: '-8,0%' },
      { k: 'Investimento', v: '3.826 €', d: '+10,6%' },
      { k: 'CPC medio', v: '0,30 €', d: '+18,5%' },
      { k: 'Quota traffico paid', v: '42,4%', d: '' }
    ]
  }
];

paidPlatforms.forEach((p, i) => {
  const x = 0.5 + i * 6.3;
  s7.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.7, w: 5.9, h: 0.6, rectRadius: 0.08, fill: { color: p.color }
  });
  s7.addText(p.name, {
    x: x + 0.2, y: 1.75, w: 5, h: 0.5,
    fontSize: 16, fontFace: FONT, bold: true, color: WHITE, isTextBox: true, margin: 0
  });
  s7.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 2.35, w: 5.9, h: 3.75, rectRadius: 0.08, fill: { color: LGRAY }
  });
  p.metrics.forEach((m, j) => {
    const y = 2.55 + j * 0.72;
    s7.addText(m.k.toUpperCase(), {
      x: x + 0.25, y: y, w: 3, h: 0.2,
      fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0
    });
    s7.addText(m.v, {
      x: x + 0.25, y: y + 0.2, w: 2.5, h: 0.4,
      fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
    });
    if (m.d) {
      const isPositive = m.d.startsWith('+');
      const isCPC = m.k.toLowerCase().includes('cpc');
      let deltaColor = isPositive ? GREEN : RED;
      if (isCPC) deltaColor = isPositive ? RED : GREEN;
      s7.addText(m.d, {
        x: x + 3.2, y: y + 0.2, w: 2.2, h: 0.4,
        fontSize: 14, fontFace: FONT, bold: true, color: deltaColor, align: 'right', isTextBox: true, margin: 0
      });
    }
  });
});

s7.addText('Investimento totale settembre 2026: 9.019 € (Google 57,6% + Meta 42,4%)', {
  x: 0.5, y: 6.3, w: 12, h: 0.4,
  fontSize: 11, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});
addFooter(s7, 'Fonte dati: Google Ads, Meta Ads | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 8 - GOOGLE ADS | DETTAGLIO
// =============================================
const s8 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s8, 'GOOGLE ADS | DETTAGLIO CAMPAGNE');
addTitle(s8, 'Demand Gen e Search Generic sono i principali driver di crescita');

const gadsHeader = [[
  { text: 'Campagna', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Clic', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Impression', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'CTR', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Costo', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'CPC', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const gadsData = [
  ['Search Generic (Consid.)', '17.384', '92.951', '18,7%', '2.827 €', '0,16 €'],
  ['DG Consideration Traffico', '19.737', '295.383', '6,7%', '1.251 €', '0,06 €'],
  ['DG Conversion Lead Gen', '2.288', '64.777', '3,5%', '636 €', '0,28 €'],
  ['Search Brand Product', '1.911', '5.122', '37,3%', '240 €', '0,13 €'],
  ['Search Brand Protection', '2.911', '6.804', '42,8%', '239 €', '0,08 €']
];
const gadsTotals = [[
  { text: 'Totale', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, fill: { color: LGRAY } } },
  { text: '44.231', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '465.037', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '9,5%', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '5.193 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '0,12 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } }
]];

const gadsRows = [
  ...gadsHeader,
  ...gadsData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: { fontSize: 10, fontFace: FONT, color: DARK, align: j > 0 ? 'right' : 'left', fill: { color: i % 2 === 0 ? WHITE : LGRAY } }
  }))),
  ...gadsTotals
];

s8.addTable(gadsRows, {
  x: 0.5, y: 1.7, w: 12.3,
  border: { type: 'none' },
  colW: [3.5, 1.5, 1.5, 1.2, 1.5, 1.2],
  rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.4],
  autoPage: false
});

s8.addText('DISTRIBUZIONE CLIC PER CAMPAGNA', {
  x: 0.5, y: 4.5, w: 5, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

s8.addChart(pres.charts.BAR, [
  { name: 'Clic', labels: ['DG Traffico', 'Search Generic', 'Brand Prot.', 'DG Lead Gen', 'Brand Product'], values: [19737, 17384, 2911, 2288, 1911] }
], {
  x: 0.5, y: 4.75, w: 7.5, h: 1.8,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [RED],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 8, dataLabelFontFace: '+mn-lt',
  showLegend: false,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 8, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin'
});

s8.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y: 4.5, w: 4.3, h: 2.1, rectRadius: 0.08, fill: { color: LGRAY } });
s8.addText('YOY HIGHLIGHTS', { x: 8.7, y: 4.6, w: 3.5, h: 0.3, fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0 });
s8.addText([
  { text: 'Clic +465%', options: { bold: true, breakLine: true, fontSize: 11 } },
  { text: 'Investimento +102% (da 2.573 a 5.193 €)', options: { breakLine: true, fontSize: 10 } },
  { text: 'CPC -64% (da 0,33 a 0,12 €)', options: { breakLine: true, fontSize: 10 } },
  { text: 'Demand Gen non era attivo nel 2025', options: { fontSize: 10, italic: true } }
], {
  x: 8.7, y: 4.95, w: 3.9, h: 1.5,
  fontFace: FONT, color: DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.4
});
addFooter(s8, 'Fonte dati: Google Ads | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 9 - META ADS | DETTAGLIO
// =============================================
const s9 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s9, 'META ADS | DETTAGLIO CAMPAGNE');
addTitle(s9, 'Blog e Lead Gen rivenditori concentrano la maggior parte dell\'investimento');

const metaHeader = [[
  { text: 'Campagna', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Link clic', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'CTR', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'CPC', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Spesa', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const metaData = [
  ['Lead Gen Rivenditori DEF', '2.308', '1,92%', '0,48 €', '1.108 €'],
  ['Blog Consideration', '7.059', '2,26%', '0,15 €', '1.061 €'],
  ['FEEL5 B2B', '575', '0,92%', '0,99 €', '572 €'],
  ['Iscrizioni Area Pro B2B', '491', '0,87%', '1,13 €', '556 €'],
  ['Promo PED', '429', '0,07%', '0,68 €', '292 €'],
  ['Store Locator B2C', '2.108', '2,01%', '0,11 €', '237 €']
];
const metaTotals = [[
  { text: 'Totale', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, fill: { color: LGRAY } } },
  { text: '12.970', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '1,01%', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '0,30 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '3.826 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } }
]];

const metaRows = [
  ...metaHeader,
  ...metaData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: { fontSize: 10, fontFace: FONT, color: DARK, align: j > 0 ? 'right' : 'left', fill: { color: i % 2 === 0 ? WHITE : LGRAY } }
  }))),
  ...metaTotals
];

s9.addTable(metaRows, {
  x: 0.5, y: 1.7, w: 12.3,
  border: { type: 'none' },
  colW: [3.5, 1.5, 1.2, 1.2, 1.5],
  rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.4],
  autoPage: false
});

s9.addText('CPC PER CAMPAGNA (€)', {
  x: 0.5, y: 4.4, w: 5, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: BLUE, isTextBox: true, margin: 0
});

s9.addChart(pres.charts.BAR, [
  { name: 'CPC', labels: ['Store Locator', 'Blog Consid.', 'Lead Gen Riv.', 'Promo PED', 'FEEL5 B2B', 'Area Pro B2B'], values: [0.11, 0.15, 0.48, 0.68, 0.99, 1.13] }
], {
  x: 0.5, y: 4.65, w: 7.5, h: 2.0,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [BLUE],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 9, dataLabelFontFace: '+mn-lt',
  showLegend: false,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 8, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  catAxisOrientation: 'minMax'
});

s9.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y: 4.4, w: 4.3, h: 2.2, rectRadius: 0.08, fill: { color: LGRAY } });
s9.addText('INSIGHT', { x: 8.7, y: 4.5, w: 3.5, h: 0.3, fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0 });
s9.addText([
  { text: 'CPC: spread 10x tra le campagne', options: { bold: true, breakLine: true, fontSize: 11 } },
  { text: 'Store Locator (0,11 €) e Blog (0,15 €) sono i piu efficienti per clic', options: { breakLine: true, fontSize: 10 } },
  { text: '\nLead Gen e B2B hanno CPC piu alti', options: { bold: true, breakLine: true, fontSize: 11 } },
  { text: 'ma portano conversioni dirette (lead, registrazioni)', options: { fontSize: 10 } }
], {
  x: 8.7, y: 4.85, w: 3.9, h: 1.6,
  fontFace: FONT, color: DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3
});
addFooter(s9, 'Fonte dati: Meta Ads | Periodo: set 2026');

// =============================================
// SLIDE 9B - QUALITA TRAFFICO PER CANALE
// =============================================
const s9b = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s9b, 'TRAFFICO | QUALITA PER CANALE');
addTitle(s9b, 'Paid Search domina per engagement (95,6%), Organic Search per profondita di navigazione (4,8 pagine)');

s9b.addText('ENGAGEMENT RATE PER CANALE (%)', {
  x: 0.5, y: 1.65, w: 6, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

s9b.addChart(pres.charts.BAR, [
  { name: 'Engagement Rate %', labels: ['Paid Search', 'Cross-network', 'Organic Social', 'Paid Social', 'Organic Search', 'Direct', 'Referral', 'Display'], values: [95.6, 89.1, 88.9, 82.6, 81.6, 78.9, 77.4, 60.3] }
], {
  x: 0.2, y: 1.9, w: 6.8, h: 3.6,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [RED],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 9, dataLabelFontFace: '+mn-lt',
  showLegend: false,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 9, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin',
  valAxisMaxVal: 110
});

// Right side table
s9b.addText('DETTAGLIO METRICHE DI QUALITA', {
  x: 7.3, y: 1.65, w: 5.5, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const qualHeader = [[
  { text: 'Canale', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Sessioni', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Pag/sess', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Durata', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const qualData = [
  ['Paid Search', '8.129', '4,73', '12m 09s'],
  ['Paid Social', '5.933', '1,62', '4m 09s'],
  ['Organic Search', '4.271', '4,80', '31m 39s'],
  ['Cross-network', '3.917', '1,33', '1m 57s'],
  ['Direct', '1.914', '3,66', '1h 48m'],
  ['Referral', '239', '3,97', '34m 19s'],
  ['Organic Social', '234', '2,59', '18m 37s'],
  ['Display', '116', '1,28', '1m 24s']
];
const qualRows = [
  ...qualHeader,
  ...qualData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: { fontSize: 8, fontFace: FONT, color: DARK, align: j > 0 ? 'right' : 'left', fill: { color: i % 2 === 0 ? WHITE : LGRAY } }
  })))
];

s9b.addTable(qualRows, {
  x: 7.3, y: 1.95, w: 5.5,
  border: { type: 'none' },
  colW: [1.6, 0.9, 0.9, 1.2],
  rowH: [0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3],
  autoPage: false
});

// Bottom insight bar
s9b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.8, w: 12.3, h: 0.9, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s9b.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.8, w: 0.06, h: 0.9, fill: { color: RED } });
s9b.addText([
  { text: 'Paid Search porta il miglior engagement (95,6%) con 4,73 pagine per sessione. ', options: { fontSize: 10 } },
  { text: 'Organic Search eccelle in profondita (4,80 pag.) e durata (31 min). ', options: { fontSize: 10, bold: true } },
  { text: 'Paid Social ha volumi elevati (5.933 sessioni) ma engagement e profondita inferiori, segnalando traffico orientato all\'atterraggio sulla landing page.', options: { fontSize: 10 } }
], {
  x: 0.8, y: 5.85, w: 11.7, h: 0.8,
  fontFace: FONT, color: DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15
});
addFooter(s9b, 'Fonte dati: GA4 | Periodo: set 2026');

// =============================================
// SLIDE 10 - EFFICIENZA CROSS-PLATFORM
// =============================================
const s10 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addSlideBase(s10, 'ADVERTISING | EFFICIENZA CROSS-PLATFORM');
addTitle(s10, 'Meta genera lead a 23,8 € (3,5x piu economico), Google porta traffico con engagement del 95,6%');

// LEFT - CPL comparison
s10.addText('COSTO PER LEAD (GA4 ATTRIBUTION)', {
  x: 0.5, y: 1.65, w: 6, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

// Meta CPL card
s10.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.0, w: 5.8, h: 1.8, rectRadius: 0.08, fill: { color: LGRAY } });
s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.0, w: 0.06, h: 1.8, fill: { color: BLUE } });
s10.addText('META ADS', { x: 0.8, y: 2.1, w: 2, h: 0.25, fontSize: 10, fontFace: FONT, bold: true, color: BLUE, isTextBox: true, margin: 0 });
s10.addText('23,8 €', { x: 0.8, y: 2.4, w: 2.5, h: 0.6, fontSize: 36, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
s10.addText('COSTO PER LEAD', { x: 0.8, y: 3.0, w: 2, h: 0.2, fontSize: 8, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
const metaCplKpis = [
  { label: 'LEAD', val: '161', kx: 3.4 },
  { label: 'SPESA', val: '3.826 €', kx: 4.5 },
  { label: '% LEAD TOT.', val: '55,3%', kx: 5.6 }
];
metaCplKpis.forEach(kpi => {
  s10.addText(kpi.label, { x: kpi.kx, y: 2.15, w: 1.1, h: 0.15, fontSize: 7, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s10.addText(kpi.val, { x: kpi.kx, y: 2.35, w: 1.1, h: 0.4, fontSize: 16, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
});
s10.addText('Best performer per costo: Blog Consideration (CPC 0,15 €) e Store Locator (CPC 0,11 €) abbassano il CPL medio', {
  x: 0.8, y: 3.25, w: 5.3, h: 0.35,
  fontSize: 8, fontFace: FONT, italic: true, color: GRAY, isTextBox: true, margin: 0
});

// Google CPL card
s10.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.0, w: 5.8, h: 1.8, rectRadius: 0.08, fill: { color: LGRAY } });
s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 0.06, h: 1.8, fill: { color: RED } });
s10.addText('GOOGLE ADS', { x: 0.8, y: 4.1, w: 2, h: 0.25, fontSize: 10, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0 });
s10.addText('82,4 €', { x: 0.8, y: 4.4, w: 2.5, h: 0.6, fontSize: 36, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
s10.addText('COSTO PER LEAD', { x: 0.8, y: 5.0, w: 2, h: 0.2, fontSize: 8, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
const googleCplKpis = [
  { label: 'LEAD', val: '63', kx: 3.4 },
  { label: 'SPESA', val: '5.193 €', kx: 4.5 },
  { label: '% LEAD TOT.', val: '21,6%', kx: 5.6 }
];
googleCplKpis.forEach(kpi => {
  s10.addText(kpi.label, { x: kpi.kx, y: 4.15, w: 1.1, h: 0.15, fontSize: 7, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s10.addText(kpi.val, { x: kpi.kx, y: 4.35, w: 1.1, h: 0.4, fontSize: 16, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
});
s10.addText('CPL piu alto ma traffico di qualita superiore: Demand Gen ha il CPC piu basso (0,06 €) e genera volumi elevati', {
  x: 0.8, y: 5.25, w: 5.3, h: 0.35,
  fontSize: 8, fontFace: FONT, italic: true, color: GRAY, isTextBox: true, margin: 0
});

// RIGHT - Quality comparison
s10.addText('QUALITA DEL TRAFFICO A CONFRONTO', {
  x: 6.8, y: 1.65, w: 6, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const effComparisons = [
  { metric: 'ENGAGEMENT RATE', google: '95,6%', meta: '82,6%', winner: 'google' },
  { metric: 'PAGINE / SESSIONE', google: '4,73', meta: '1,62', winner: 'google' },
  { metric: 'DURATA MEDIA SESSIONE', google: '12m 09s', meta: '4m 09s', winner: 'google' },
  { metric: 'CLIC TOTALI', google: '44.231', meta: '12.970', winner: 'google' },
  { metric: 'COSTO PER LEAD', google: '82,4 €', meta: '23,8 €', winner: 'meta' }
];

effComparisons.forEach((comp, i) => {
  const y = 2.0 + i * 0.95;
  s10.addText(comp.metric, {
    x: 6.8, y: y, w: 6, h: 0.2,
    fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0
  });
  s10.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: y + 0.25, w: 3.0, h: 0.55, rectRadius: 0.06, fill: { color: comp.winner === 'google' ? 'FFE7E6' : LGRAY } });
  s10.addText('Google', { x: 6.9, y: y + 0.27, w: 1.2, h: 0.15, fontSize: 7, fontFace: FONT, color: RED, isTextBox: true, margin: 0 });
  s10.addText(comp.google, { x: 6.9, y: y + 0.42, w: 2.8, h: 0.3, fontSize: 16, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s10.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.0, y: y + 0.25, w: 2.8, h: 0.55, rectRadius: 0.06, fill: { color: comp.winner === 'meta' ? 'E6EEFF' : LGRAY } });
  s10.addText('Meta', { x: 10.1, y: y + 0.27, w: 1.2, h: 0.15, fontSize: 7, fontFace: FONT, color: BLUE, isTextBox: true, margin: 0 });
  s10.addText(comp.meta, { x: 10.1, y: y + 0.42, w: 2.6, h: 0.3, fontSize: 16, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
});

// Bottom insight
s10.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 6.05, w: 12.3, h: 0.7, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 6.05, w: 0.06, h: 0.7, fill: { color: RED } });
s10.addText([
  { text: 'Meta e Google svolgono ruoli complementari: ', options: { fontSize: 10, bold: true } },
  { text: 'Meta e 3,5x piu efficiente per lead generation (55% delle lead a costo inferiore), mentre Google Ads porta utenti con il 95,6% di engagement e quasi 5 pagine per sessione, alimentando tutto il funnel.', options: { fontSize: 10 } }
], {
  x: 0.8, y: 6.1, w: 11.7, h: 0.6,
  fontFace: FONT, color: DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15
});
addFooter(s10, 'Fonte dati: GA4, Google Ads, Meta Ads | Periodo: set 2026');

// =============================================
// SLIDE 11 - DIVIDER: LEAD GENERATION
// =============================================
pres.addSection({ title: 'Lead Generation' });
const s11 = pres.addSlide({ sectionTitle: 'Lead Generation' });
s11.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s11.addText([
  { text: 'Lead Generation', options: { fontSize: 38, fontFace: FONT, color: WHITE, italic: true, breakLine: true } },
  { text: 'Settembre 2026.', options: { fontSize: 38, fontFace: FONT, color: '888888', italic: true } }
], { x: 1.0, y: 2.5, w: 11, h: 2.5, isTextBox: true, margin: 0 });

// =============================================
// SLIDE 12 - LEAD GENERATION | PANORAMICA (bar chart + KPIs)
// =============================================
const s12 = pres.addSlide({ sectionTitle: 'Lead Generation' });
addSlideBase(s12, 'LEAD GENERATION | PANORAMICA');
addTitle(s12, '459 lead totali: la landing rivenditori porta 215 preventivi, compensando il calo tradizionale');

const leadKpis = [
  { label: 'LEAD TOTALI', val: '459', sub: 'settembre 2026', color: DARK },
  { label: 'PREVENTIVI LP', val: '215', sub: 'nuovo canale', color: RED },
  { label: 'LEAD TRADIZIONALI', val: '244', sub: '-24,5% vs 2025', color: RED }
];
leadKpis.forEach((kpi, i) => {
  const x = 0.5 + i * 4.2;
  s12.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: 1.7, w: 3.9, h: 1.3, rectRadius: 0.08, fill: { color: LGRAY } });
  s12.addText(kpi.label, { x: x + 0.2, y: 1.8, w: 3.5, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s12.addText(kpi.val, { x: x + 0.2, y: 2.05, w: 2, h: 0.6, fontSize: 36, fontFace: FONT, bold: true, color: kpi.color, isTextBox: true, margin: 0 });
  s12.addText(kpi.sub, { x: x + 0.2, y: 2.6, w: 3, h: 0.25, fontSize: 9, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
});

s12.addChart(pres.charts.BAR, [
  { name: 'Set 2026', labels: ['Info prodotto', 'Registrazioni', 'FEEL5', 'Form contatto', 'Contatta PV', 'FEEL4', 'Newsletter'], values: [61, 56, 54, 38, 25, 7, 3] },
  { name: 'Set 2025', labels: ['Info prodotto', 'Registrazioni', 'FEEL5', 'Form contatto', 'Contatta PV', 'FEEL4', 'Newsletter'], values: [100, 55, 99, 36, 29, 2, 2] }
], {
  x: 0.5, y: 3.2, w: 12.3, h: 3.2,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [RED, 'CCCCCC'],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 8, dataLabelFontFace: '+mn-lt',
  showLegend: true, legendPos: 't', legendFontSize: 9, legendFontFace: '+mn-lt', legendColor: GRAY,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 9, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin'
});
addFooter(s12, 'Fonte dati: HubSpot, MyArea | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 12B - LEAD PER CANALE PAID
// =============================================
const s12b = pres.addSlide({ sectionTitle: 'Lead Generation' });
addSlideBase(s12b, 'LEAD GENERATION | ATTRIBUZIONE PER CANALE');
addTitle(s12b, 'Meta genera il 55% delle lead tracciate con un costo per lead inferiore del 60% rispetto a Google');

s12b.addChart(pres.charts.DOUGHNUT, [
  { name: 'Lead per canale', labels: ['Meta Ads', 'Google Ads', 'Organico', 'Diretto e altro'], values: [161, 63, 34, 33] }
], {
  x: 0.0, y: 1.7, w: 5.5, h: 4.0,
  showTitle: false,
  chartColors: [BLUE, RED, GREEN, DGRAY],
  showLegend: true, legendPos: 'b', legendFontSize: 9, legendFontFace: '+mn-lt', legendColor: DARK,
  showPercent: true,
  dataLabelPosition: 'ctr',
  dataLabelFontSize: 13,
  dataLabelFontFace: '+mn-lt',
  dataLabelColor: WHITE
});

const leadChHeader = [[
  { text: 'Canale', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Lead 2026', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: '% totale', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Lead 2025', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Var. lead', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Costo 2026', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Var. costo', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const leadChData = [
  ['Meta Ads', '161', '55,3%', '74', '+117,6%', '3.826 €', '+10,6%'],
  ['Google Ads', '63', '21,6%', '48', '+31,3%', '5.193 €', '+101,7%'],
  ['Organico', '34', '11,7%', '52', '-34,6%', '-', '-'],
  ['Diretto e altro', '33', '11,3%', '27', '+22,2%', '-', '-']
];
const leadChTotals = [[
  { text: 'Totale', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, fill: { color: LGRAY } } },
  { text: '291', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '100%', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '201', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '+44,8%', options: { bold: true, fontSize: 10, fontFace: FONT, color: GREEN, align: 'right', fill: { color: LGRAY } } },
  { text: '9.019 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '+48,3%', options: { bold: true, fontSize: 10, fontFace: FONT, color: RED, align: 'right', fill: { color: LGRAY } } }
]];

const leadChRows = [
  ...leadChHeader,
  ...leadChData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: {
      fontSize: 10, fontFace: FONT,
      color: (j === 4 || j === 6) && cell.startsWith('+') ? GREEN : (j === 4 || j === 6) && cell.startsWith('-') ? RED : DARK,
      align: j > 0 ? 'right' : 'left',
      fill: { color: i % 2 === 0 ? WHITE : LGRAY }
    }
  }))),
  ...leadChTotals
];

s12b.addTable(leadChRows, {
  x: 5.5, y: 1.7, w: 7.3,
  border: { type: 'none' },
  colW: [1.4, 0.9, 0.8, 0.9, 0.9, 1.0, 0.9],
  rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.4],
  autoPage: false
});

s12b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.9, w: 12.3, h: 0.8, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s12b.addText([
  { text: 'Meta Ads: CPL 23,8 € (161 lead con 3.826 €) ', options: { bold: true, color: RED } },
  { text: 'vs ', options: { color: GRAY } },
  { text: 'Google Ads: CPL 82,4 € (63 lead con 5.193 €). ', options: { bold: true, color: RED } },
  { text: 'Meta porta il 55% delle lead con il 42% del budget, confermandosi il canale paid piu efficiente per la generazione di contatti.', options: { color: GRAY } }
], {
  x: 0.7, y: 5.95, w: 11.8, h: 0.7,
  fontSize: 10, fontFace: FONT, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3
});
addFooter(s12b, 'Fonte dati: GA4 (lead event), Google Ads, Meta Ads | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 13 - LEAD GENERATION | DETTAGLIO (table)
// =============================================
const s13 = pres.addSlide({ sectionTitle: 'Lead Generation' });
addSlideBase(s13, 'LEAD GENERATION | DETTAGLIO');
addTitle(s13, 'Dettaglio completo per tipologia di lead e fonte dati');

const leadHeader = [[
  { text: 'Tipologia lead', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Fonte', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Set 2026', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Set 2025', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Var. %', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const leadDataRows = [
  ['Preventivo LP (landing rivenditori)', 'HubSpot', '215', '-', 'nuovo'],
  ['Info prodotto (totale)', 'HubSpot + MyArea', '61', '100', '-39,0%'],
  ['Registrazioni area', 'HubSpot', '56', '55', '+1,8%'],
  ['FEEL5', 'HubSpot', '54', '99', '-45,5%'],
  ['Form contatto (IT/EN/ES/FR)', 'HubSpot', '38', '36', '+5,6%'],
  ['Contatta PV', 'MyArea', '25', '29', '-13,8%'],
  ['FEEL4', 'HubSpot', '7', '2', '+250,0%'],
  ['Newsletter', 'HubSpot', '3', '2', '+50,0%']
];
const leadTot1 = [[
  { text: 'Totale (escl. Preventivo LP)', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, fill: { color: 'FFF0EF' } } },
  { text: '', options: { fontSize: 10, fontFace: FONT, fill: { color: 'FFF0EF' } } },
  { text: '244', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: 'FFF0EF' } } },
  { text: '323', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: 'FFF0EF' } } },
  { text: '-24,5%', options: { bold: true, fontSize: 10, fontFace: FONT, color: RED, align: 'right', fill: { color: 'FFF0EF' } } }
]];
const leadTot2 = [[
  { text: 'Totale (incl. Preventivo LP)', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, fill: { color: LGRAY } } },
  { text: '', options: { fontSize: 10, fontFace: FONT, fill: { color: LGRAY } } },
  { text: '459', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '', options: { fontSize: 10, fontFace: FONT, fill: { color: LGRAY } } },
  { text: '', options: { fontSize: 10, fontFace: FONT, fill: { color: LGRAY } } }
]];

const leadRows = [
  ...leadHeader,
  ...leadDataRows.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: {
      fontSize: 10, fontFace: FONT,
      color: j === 4 && cell.startsWith('+') ? GREEN : j === 4 && cell.startsWith('-') ? RED : DARK,
      align: j > 1 ? 'right' : 'left',
      fill: { color: i % 2 === 0 ? WHITE : LGRAY }
    }
  }))),
  ...leadTot1,
  ...leadTot2
];

s13.addTable(leadRows, {
  x: 0.5, y: 1.7, w: 12.3,
  border: { type: 'none' },
  colW: [4.0, 2.0, 1.8, 1.8, 1.8],
  rowH: [0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.45, 0.45],
  autoPage: false
});

s13.addText([
  { text: 'La nuova landing rivenditori (Preventivo LP) ', options: { bold: true } },
  { text: 'è il dato più significativo del mese: 215 preventivi da un canale completamente nuovo. Tra le lead tradizionali, il calo di FEEL5 (-45,5%) e info prodotto (-39%) è parzialmente compensato dalla stabilità delle registrazioni area (+1,8%) e dei form contatto (+5,6%).' }
], {
  x: 0.5, y: 6.0, w: 12.3, h: 0.8,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});
addFooter(s13, 'Fonte dati: HubSpot, MyArea | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 14 - LEAD GENERATION | FOCUS RIVENDITORI
// =============================================
const s14 = pres.addSlide({ sectionTitle: 'Lead Generation' });
addSlideBase(s14, 'LEAD GENERATION | FOCUS RIVENDITORI');
addTitle(s14, '215 preventivi dalla landing rivenditori: Meta porta 120 lead, Google 26');

s14.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.7, w: 5.9, h: 4.7, rectRadius: 0.1, fill: { color: LGRAY } });
s14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 5.9, h: 0.08, fill: { color: RED } });
s14.addText('Campagna Google Ads', { x: 0.8, y: 1.9, w: 5, h: 0.35, fontSize: 14, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
s14.addText('DG Conversion Lead Gen', { x: 0.8, y: 2.25, w: 5, h: 0.25, fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });

const gLeadKpis = [
  { k: 'IMPRESSION', v: '64.777' },
  { k: 'CLIC', v: '2.288' },
  { k: 'INVESTIMENTO', v: '636 €' },
  { k: 'CPC', v: '0,28 €' }
];
gLeadKpis.forEach((kpi, i) => {
  const y = 2.8 + i * 0.55;
  s14.addText(kpi.k, { x: 0.8, y: y, w: 1.8, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s14.addText(kpi.v, { x: 2.8, y: y - 0.02, w: 2.5, h: 0.28, fontSize: 15, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
});

s14.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 5.0, w: 5.3, h: 0.9, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s14.addText([
  { text: '~26 ', options: { fontSize: 24, bold: true, color: RED } },
  { text: 'lead attribuite (GA4)', options: { fontSize: 11, color: DARK } }
], { x: 1.0, y: 5.0, w: 4.8, h: 0.45, fontFace: FONT, isTextBox: true, margin: 0 });
s14.addText('Paid Search 9 + Cross-network 17', {
  x: 1.0, y: 5.45, w: 4.8, h: 0.3, fontSize: 9, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

s14.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: 1.7, w: 5.9, h: 4.7, rectRadius: 0.1, fill: { color: LGRAY } });
s14.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 1.7, w: 5.9, h: 0.08, fill: { color: BLUE } });
s14.addText('Campagna Meta Ads', { x: 7.1, y: 1.9, w: 5, h: 0.35, fontSize: 14, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
s14.addText('Lead Gen Rivenditori DEF', { x: 7.1, y: 2.25, w: 5, h: 0.25, fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });

const mLeadKpis = [
  { k: 'REACH', v: '36.082' },
  { k: 'LINK CLIC', v: '2.308' },
  { k: 'INVESTIMENTO', v: '1.108 €' },
  { k: 'CTR', v: '1,92%' }
];
mLeadKpis.forEach((kpi, i) => {
  const y = 2.8 + i * 0.55;
  s14.addText(kpi.k, { x: 7.1, y: y, w: 1.8, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s14.addText(kpi.v, { x: 9.1, y: y - 0.02, w: 2.5, h: 0.28, fontSize: 15, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
});

s14.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.1, y: 5.0, w: 5.3, h: 0.9, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s14.addText([
  { text: '120 ', options: { fontSize: 24, bold: true, color: RED } },
  { text: 'lead attribuite (GA4)', options: { fontSize: 11, color: DARK } }
], { x: 7.3, y: 5.0, w: 4.8, h: 0.45, fontFace: FONT, isTextBox: true, margin: 0 });
s14.addText('Paid Social, submit_contact_form_lp', {
  x: 7.3, y: 5.45, w: 4.8, h: 0.3, fontSize: 9, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

s14.addText('Il totale HubSpot (215 preventivi) supera le lead tracciate in GA4 perche include compilazioni da utenti di ritorno, organico e diretto non sempre attribuibili ai canali paid.', {
  x: 0.5, y: 6.2, w: 12.3, h: 0.5,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});
addFooter(s14, 'Fonte dati: GA4, HubSpot, Google Ads, Meta Ads | Periodo: set 2026');

// =============================================
// SLIDE 15 - DIVIDER: FOCUS MAGAZINE
// =============================================
pres.addSection({ title: 'Focus Magazine' });
const s15 = pres.addSlide({ sectionTitle: 'Focus Magazine' });
s15.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s15.addText([
  { text: 'Focus Magazine.', options: { fontSize: 38, fontFace: FONT, color: WHITE, italic: true } }
], { x: 1.0, y: 2.8, w: 11, h: 2, isTextBox: true, margin: 0 });

// =============================================
// SLIDE 16 - MAGAZINE | ORGANICO (big % KPIs)
// =============================================
const s16 = pres.addSlide({ sectionTitle: 'Focus Magazine' });
addSlideBase(s16, 'TRAFFICO ORGANICO | MAGAZINE');
addTitle(s16, 'Il magazine guadagna 10 posizioni ma perde clic per effetto AI Overview');

const magMetrics = [
  { pct: '-21%', label: 'CLIC', val: '92 vs 117', color: RED },
  { pct: '-17%', label: 'IMPRESSION', val: '8.267 vs 9.969', color: RED },
  { pct: '1,11%', label: 'CTR MEDIO', val: 'vs 1,17%', color: GRAY },
  { pct: '8,7', label: 'POSIZIONE MEDIA', val: 'vs 19,4 (+10,7 pos.)', color: GREEN }
];

magMetrics.forEach((m, i) => {
  const x = 0.5 + i * 3.1;
  s16.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.7, w: 2.8, h: 3.0, rectRadius: 0.1, fill: { color: LGRAY }
  });
  s16.addText(m.pct, {
    x: x + 0.2, y: 1.9, w: 2.4, h: 1.2,
    fontSize: 42, fontFace: FONT, bold: true, color: m.color, isTextBox: true, margin: 0
  });
  s16.addText(m.label, {
    x: x + 0.2, y: 3.1, w: 2.4, h: 0.3,
    fontSize: 10, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s16.addText(m.val, {
    x: x + 0.2, y: 3.45, w: 2.4, h: 0.3,
    fontSize: 9, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
  });
});

s16.addText([
  { text: 'La posizione media migliora da 19,4 a 8,7 (+10,7 posizioni): ', options: { bold: true } },
  { text: 'un risultato molto significativo. Tuttavia i clic diminuiscono del 21% per la contrazione delle impression (-17%). L\'erosione del CTR da AI Overview continua a incidere particolarmente sui contenuti informativi del magazine.' }
], {
  x: 0.5, y: 5.0, w: 12.3, h: 0.8,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

s16.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.9, w: 12.3, h: 0.8, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s16.addText('Download magazine FEEL: 35 (set 2026) vs 69 (set 2025) = -49%', {
  x: 0.8, y: 5.95, w: 11, h: 0.6,
  fontSize: 12, fontFace: FONT, bold: true, color: RED, valign: 'middle', isTextBox: true, margin: 0
});
addFooter(s16, 'Fonte dati: GSC, GA4 | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 17 - MAGAZINE | PAID
// =============================================
const s17 = pres.addSlide({ sectionTitle: 'Focus Magazine' });
addSlideBase(s17, 'TRAFFICO PAID | MAGAZINE');
addTitle(s17, 'La campagna Blog Consideration porta 7.059 clic al magazine con CPC molto efficiente');

s17.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.7, w: 12.3, h: 2.0, rectRadius: 0.1, fill: { color: LGRAY } });
s17.addText('BLOG CONSIDERATION (META ADS)', { x: 0.8, y: 1.85, w: 8, h: 0.3, fontSize: 10, fontFace: FONT, bold: true, color: BLUE, isTextBox: true, margin: 0 });

const blogKpis = [
  { label: 'LINK CLIC', val: '7.059', x: 0.8 },
  { label: 'REACH', val: '136.533', x: 3.2 },
  { label: 'CTR', val: '2,26%', x: 5.6 },
  { label: 'CPC MEDIO', val: '0,15 €', x: 8.0 },
  { label: 'INVESTIMENTO', val: '1.061 €', x: 10.5 }
];
blogKpis.forEach(kpi => {
  s17.addText(kpi.label, { x: kpi.x, y: 2.3, w: 2.3, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s17.addText(kpi.val, { x: kpi.x, y: 2.55, w: 2.3, h: 0.5, fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
});

s17.addText('SESSIONI ALLE PAGINE MAGAZINE (GA4)', { x: 0.5, y: 4.0, w: 8, h: 0.3, fontSize: 10, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });

const magPages = [
  { page: 'Nuovo magazine FEEL5 (landing)', sessions: 445 },
  { page: 'Magazine FEEL (indice)', sessions: 117 },
  { page: 'Download magazine FEEL5', sessions: 37 },
  { page: 'Magazine FEEL4 (landing)', sessions: 26 }
];

s17.addChart(pres.charts.BAR, [
  { name: 'Sessioni', labels: magPages.map(p => p.page), values: magPages.map(p => p.sessions) }
], {
  x: 0.5, y: 4.3, w: 12.3, h: 2.2,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [RED],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 10, dataLabelFontFace: '+mn-lt',
  showLegend: false,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 9, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin'
});
addFooter(s17, 'Fonte dati: Meta Ads, GA4 | Periodo: set 2026');

// =============================================
// SLIDE 17B - PROMO PED E CANALI SOCIAL
// =============================================
const s17b = pres.addSlide({ sectionTitle: 'Focus Magazine' });
addSlideBase(s17b, 'SOCIAL | PROMO PED E CANALI ORGANICI');
addTitle(s17b, 'Promo PED raggiunge 469K utenti unici con 292 €, i social organici portano 234 sessioni al sito');

// LEFT - Promo PED
s17b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.7, w: 6.2, h: 4.8, rectRadius: 0.1, fill: { color: LGRAY } });
s17b.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.7, w: 0.06, h: 4.8, fill: { color: BLUE } });
s17b.addText('PROMO PED (META ADS, AWARENESS)', { x: 0.8, y: 1.85, w: 5.5, h: 0.3, fontSize: 10, fontFace: FONT, bold: true, color: BLUE, isTextBox: true, margin: 0 });
s17b.addText('Campagna always-on di brand awareness su Facebook e Instagram', {
  x: 0.8, y: 2.2, w: 5.5, h: 0.3,
  fontSize: 9, fontFace: FONT, italic: true, color: GRAY, isTextBox: true, margin: 0
});

const pedKpis = [
  { label: 'IMPRESSION', val: '631.044', yoy: '+5,4%', good: true },
  { label: 'REACH', val: '468.525', yoy: '+5,9%', good: true },
  { label: 'FREQUENZA', val: '1,35', yoy: 'stabile', good: true },
  { label: 'LINK CLIC', val: '429', yoy: '+18,8%', good: true },
  { label: 'CTR', val: '0,07%', yoy: '+12,7%', good: true },
  { label: 'CPC', val: '0,68 €', yoy: '-4,8%', good: true }
];
pedKpis.forEach((kpi, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const kx = 0.8 + col * 2.0;
  const ky = 2.7 + row * 1.4;
  s17b.addText(kpi.label, { x: kx, y: ky, w: 1.8, h: 0.15, fontSize: 7, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s17b.addText(kpi.val, { x: kx, y: ky + 0.2, w: 1.8, h: 0.5, fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  const yoyColor = kpi.yoy === 'stabile' ? GRAY : GREEN;
  s17b.addText(kpi.yoy + ' YoY', { x: kx, y: ky + 0.7, w: 1.8, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: yoyColor, isTextBox: true, margin: 0 });
});

s17b.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.55, w: 5.5, h: 0.03, fill: { color: 'CCCCCC' } });
s17b.addText('INVESTIMENTO: 292 € (+13,2% YoY) | 6 creative aggiornate nel mese (album, video, carousel, immagine)', {
  x: 0.8, y: 5.7, w: 5.5, h: 0.35,
  fontSize: 9, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});
s17b.addText('CPM di 0,46 €: tra le campagne piu efficienti per copertura', {
  x: 0.8, y: 6.05, w: 5.5, h: 0.25,
  fontSize: 9, fontFace: FONT, italic: true, color: GRAY, isTextBox: true, margin: 0
});

// RIGHT - Social organici
s17b.addText('TRAFFICO SOCIAL ORGANICO AL SITO (GA4)', {
  x: 7.0, y: 1.65, w: 6, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const socHeader = [[
  { text: 'Canale', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, fill: { color: DARK } } },
  { text: 'Set 2026', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Set 2025', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } },
  { text: 'Var. %', options: { bold: true, color: WHITE, fontSize: 8, fontFace: FONT, align: 'right', fill: { color: DARK } } }
]];
const socData = [
  { n: 'Facebook', v26: '177', v25: '351', yoy: '-49,6%', neg: true },
  { n: 'Instagram', v26: '10', v25: '2', yoy: '+400%', neg: false },
  { n: 'Pinterest (org.)', v26: '13', v25: '32', yoy: '-59,4%', neg: true },
  { n: 'YouTube', v26: '2', v25: '0', yoy: 'nuovo', neg: false },
  { n: 'Altro (bit.ly, houzz)', v26: '34', v25: '22', yoy: '+54,5%', neg: false }
];
const socRows = socData.map((r, i) => [
  { text: r.n, options: { fontSize: 9, fontFace: FONT, color: DARK, fill: { color: i % 2 === 0 ? WHITE : LGRAY } } },
  { text: r.v26, options: { fontSize: 9, fontFace: FONT, color: DARK, align: 'right', fill: { color: i % 2 === 0 ? WHITE : LGRAY } } },
  { text: r.v25, options: { fontSize: 9, fontFace: FONT, color: DARK, align: 'right', fill: { color: i % 2 === 0 ? WHITE : LGRAY } } },
  { text: r.yoy, options: { fontSize: 9, fontFace: FONT, bold: true, color: r.neg ? RED : GREEN, align: 'right', fill: { color: i % 2 === 0 ? WHITE : LGRAY } } }
]);
const socTotal = [[
  { text: 'Totale Organic Social', options: { bold: true, fontSize: 9, fontFace: FONT, color: DARK, fill: { color: LGRAY } } },
  { text: '234', options: { bold: true, fontSize: 9, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '410', options: { bold: true, fontSize: 9, fontFace: FONT, color: DARK, align: 'right', fill: { color: LGRAY } } },
  { text: '-42,9%', options: { bold: true, fontSize: 9, fontFace: FONT, color: RED, align: 'right', fill: { color: LGRAY } } }
]];

s17b.addTable([...socHeader, ...socRows, ...socTotal], {
  x: 7.0, y: 1.95, w: 5.8,
  border: { type: 'none' },
  colW: [2.0, 1.0, 1.0, 1.0],
  rowH: [0.3, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
  autoPage: false
});

// Right side - KPI cards for Instagram and YouTube
s17b.addText('FOCUS: INSTAGRAM E YOUTUBE', {
  x: 7.0, y: 4.55, w: 6, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

s17b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 4.85, w: 2.8, h: 1.6, rectRadius: 0.06, fill: { color: 'FCE4EC' } });
s17b.addText('INSTAGRAM', { x: 7.15, y: 4.95, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: PINK, isTextBox: true, margin: 0 });
s17b.addText('10', { x: 7.15, y: 5.2, w: 2, h: 0.5, fontSize: 30, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
s17b.addText('sessioni al sito', { x: 7.15, y: 5.65, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
s17b.addText('100% engagement, 5,67 pag/sess\nCanale di awareness, non di traffico diretto', {
  x: 7.15, y: 5.9, w: 2.5, h: 0.4,
  fontSize: 7, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

s17b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.0, y: 4.85, w: 2.8, h: 1.6, rectRadius: 0.06, fill: { color: 'FFEBEE' } });
s17b.addText('YOUTUBE', { x: 10.15, y: 4.95, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0 });
s17b.addText('2', { x: 10.15, y: 5.2, w: 2, h: 0.5, fontSize: 30, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
s17b.addText('sessioni al sito', { x: 10.15, y: 5.65, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
s17b.addText('Canale emergente, 100% eng.\n8 pag/sessione, durata 2h 21m', {
  x: 10.15, y: 5.9, w: 2.5, h: 0.4,
  fontSize: 7, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

addFooter(s17b, 'Fonte dati: Meta Ads, GA4 | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 18 - PINTEREST | ANALISI APPROFONDITA
// =============================================
const s18 = pres.addSlide({ sectionTitle: 'Focus Magazine' });
addSlideBase(s18, 'PINTEREST | ANALISI APPROFONDITA');
addTitle(s18, 'Pinterest cresce del 22% in sessioni, ma l\'engagement scende di 15 punti: serve ottimizzare la qualita');

// LEFT - Hero number + bar chart
s18.addText('+22,1%', {
  x: 0.5, y: 1.8, w: 4, h: 1.0,
  fontSize: 60, fontFace: FONT, bold: true, color: GREEN, isTextBox: true, margin: 0
});
s18.addText('CRESCITA SESSIONI YOY', {
  x: 0.5, y: 2.75, w: 4, h: 0.25,
  fontSize: 10, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

s18.addChart(pres.charts.BAR, [
  { name: '', labels: ['Set 2026', 'Set 2025'], values: [1653, 1354] }
], {
  x: 0.3, y: 3.2, w: 4.2, h: 2.0,
  showTitle: false,
  barDir: 'col', barGrouping: 'clustered',
  chartColors: [PINK],
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, dataLabelFontSize: 11, dataLabelFontFace: '+mn-lt',
  showLegend: false,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 9, catAxisLabelFontFace: '+mn-lt',
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' }
});

// RIGHT - Engagement and quality metrics comparison
s18.addText('METRICHE DI QUALITA', {
  x: 5.0, y: 1.65, w: 4, h: 0.25,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const pinMetrics = [
  { metric: 'ENGAGEMENT RATE', val26: '67,6%', val25: '82,5%', delta: '-14,9 pp', negative: true },
  { metric: 'PAGINE / SESSIONE', val26: '1,52', val25: '1,71', delta: '-11,1%', negative: true },
  { metric: 'DURATA MEDIA', val26: '3m 19s', val25: '1m 47s', delta: '+85,6%', negative: false },
  { metric: 'SESSIONI TOTALI', val26: '1.653', val25: '1.354', delta: '+22,1%', negative: false }
];

pinMetrics.forEach((m, i) => {
  const y = 2.0 + i * 1.15;
  s18.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.0, y: y, w: 7.8, h: 0.95, rectRadius: 0.06, fill: { color: LGRAY } });
  s18.addText(m.metric, { x: 5.2, y: y + 0.05, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s18.addText(m.val26, { x: 5.2, y: y + 0.3, w: 2, h: 0.45, fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s18.addText('2025: ' + m.val25, { x: 7.5, y: y + 0.35, w: 2, h: 0.3, fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
  s18.addText(m.delta, { x: 10.5, y: y + 0.25, w: 2, h: 0.4, fontSize: 16, fontFace: FONT, bold: true, color: m.negative ? RED : GREEN, align: 'right', isTextBox: true, margin: 0 });
});

// COMPARISON with other channels
s18.addText('CONFRONTO CON ALTRI CANALI (ENGAGEMENT RATE)', {
  x: 5.0, y: 6.05, w: 5, h: 0.2,
  fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0
});
const pinCompare = [
  { ch: 'Paid Search', val: '95,6%', x: 5.0 },
  { ch: 'Organic Search', val: '81,6%', x: 7.2 },
  { ch: 'Paid Social', val: '82,6%', x: 9.3 },
  { ch: 'Pinterest', val: '67,6%', x: 11.3 }
];
pinCompare.forEach(c => {
  s18.addText(c.ch, { x: c.x, y: 6.3, w: 1.9, h: 0.15, fontSize: 7, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0 });
  s18.addText(c.val, { x: c.x, y: 6.45, w: 1.9, h: 0.3, fontSize: 14, fontFace: FONT, bold: true, color: c.ch === 'Pinterest' ? RED : DARK, isTextBox: true, margin: 0 });
});

// Bottom insight
s18.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.5, w: 4.2, h: 1.3, rectRadius: 0.08, fill: { color: 'FFF0EF' } });
s18.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.5, w: 0.06, h: 1.3, fill: { color: RED } });
s18.addText([
  { text: 'I volumi crescono ma la qualita cala. ', options: { bold: true, fontSize: 10, breakLine: true } },
  { text: 'Il 99% del traffico Pinterest (1.639 su 1.653) arriva da distribuzione paid. L\'engagement in calo (-15 pp) e la bassa profondita (1,52 pag.) suggeriscono di rivedere targeting e creativita.', options: { fontSize: 9 } }
], {
  x: 0.75, y: 5.6, w: 3.8, h: 1.1,
  fontFace: FONT, color: DARK, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2
});
addFooter(s18, 'Fonte dati: GA4 | Periodo: set 2026 vs set 2025');

// =============================================
// SLIDE 19 - DIVIDER: IN CHIUSURA
// =============================================
pres.addSection({ title: 'In chiusura' });
const s19 = pres.addSlide({ sectionTitle: 'In chiusura' });
s19.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s19.addText('In chiusura.', {
  x: 1.0, y: 2.8, w: 11.3, h: 2,
  fontSize: 38, fontFace: FONT, color: WHITE, italic: true, isTextBox: true, margin: 0
});

// =============================================
// SLIDE 20 - RIEPILOGO ADVERTISING (numbered list)
// =============================================
const s20 = pres.addSlide({ sectionTitle: 'In chiusura' });
addSlideBase(s20, 'RIEPILOGO | ADVERTISING');
addTitle(s20, 'I risultati chiave dell\'advertising a settembre');

const advRecap = [
  { title: 'Google Ads quadruplica i clic', text: '+465% clic (44.231 vs 7.834), con un CPC in calo del 64% grazie all\'efficienza delle campagne Demand Gen. Investimento: 5.193 € (+102%).' },
  { title: 'Demand Gen, il nuovo protagonista', text: 'Le campagne Demand Gen (assenti nel 2025) portano oltre 22.000 clic con CPC medio di 0,06 €, il canale paid più efficiente in assoluto.' },
  { title: 'Meta Ads mantiene i volumi', text: 'Blog Consideration è il best performer: 7.059 clic con CPC di 0,15 €. Lead Gen Rivenditori porta 2.308 clic con CTR dell\'1,92%.' },
  { title: 'Landing rivenditori: 215 preventivi', text: 'Il nuovo canale genera 215 preventivi tra Google Ads (DG Conversion, 636 €) e Meta Ads (Lead Gen Riv., 1.108 €).' },
  { title: 'Investimento totale: 9.019 €', text: 'Google 57,6% (5.193 €) + Meta 42,4% (3.826 €). Il paid porta il 70% del traffico totale al sito.' }
];

advRecap.forEach((item, i) => {
  const y = 1.6 + i * 1.1;
  s20.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: RED } });
  s20.addText(String(i + 1), {
    x: 0.5, y: y + 0.05, w: 0.45, h: 0.45,
    fontSize: 16, fontFace: FONT, bold: true, color: WHITE, align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });
  s20.addText(item.title, {
    x: 1.15, y: y, w: 11, h: 0.35,
    fontSize: 13, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s20.addText(item.text, {
    x: 1.15, y: y + 0.35, w: 11, h: 0.55,
    fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
  });
});
addFooter(s20, '');

// =============================================
// SLIDE 21 - RIEPILOGO SEO E ORGANICO (numbered list)
// =============================================
const s21 = pres.addSlide({ sectionTitle: 'In chiusura' });
addSlideBase(s21, 'RIEPILOGO | SEO E ORGANICO');
addTitle(s21, 'I risultati chiave del canale organico a settembre');

const seoRecap = [
  { title: 'Clic organici in calo, posizionamento in crescita', text: '-18% clic in Italia (5.391 vs 6.593), ma la posizione media migliora da 15,9 a 11,9 (+4 posizioni). Le impression crescono dell\'1,7%.' },
  { title: 'AI Overview è il principale driver del calo', text: 'Le risposte generate dall\'AI nei risultati di ricerca Google erodono il CTR dei siti web, in particolare sui contenuti informativi.' },
  { title: 'Magazine: +10 posizioni guadagnate', text: 'La posizione media del magazine migliora da 19,4 a 8,7 (+10,7 posizioni), un risultato molto significativo nonostante il calo dei clic (-21%).' },
  { title: 'Pinterest cresce del 22%', text: '1.653 sessioni da Pinterest (+22,1%), confermando il valore della presenza organica su una piattaforma allineata al target di design.' },
  { title: 'Lead tradizionali in calo', text: '244 lead tradizionali (-24,5%), in particolare FEEL5 (-45,5%) e info prodotto (-39%). Stabili registrazioni area (+1,8%) e form contatto (+5,6%).' }
];

seoRecap.forEach((item, i) => {
  const y = 1.6 + i * 1.1;
  s21.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: RED } });
  s21.addText(String(i + 1), {
    x: 0.5, y: y + 0.05, w: 0.45, h: 0.45,
    fontSize: 16, fontFace: FONT, bold: true, color: WHITE, align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });
  s21.addText(item.title, {
    x: 1.15, y: y, w: 11, h: 0.35,
    fontSize: 13, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s21.addText(item.text, {
    x: 1.15, y: y + 0.35, w: 11, h: 0.55,
    fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
  });
});
addFooter(s21, '');

// =============================================
// WRITE FILE
// =============================================
async function build() {
  await pres.writeFile({ fileName: '/home/user/ferrerolegno/FerreroLegnoReport_settembre_2026.pptx' });
  console.log('PPTX written successfully');

  const { applyTheme } = require('/root/.claude/skills/synced/413fe30c-0d0e-4c32-ae08-c0a4269ebff6_06658e08-ad09-483b-b8a9-29fe64a41207/pptx/scripts/apply_theme.js');
  await applyTheme('/home/user/ferrerolegno/FerreroLegnoReport_settembre_2026.pptx', THEME);
  console.log('Theme applied successfully');
}

build().catch(e => { console.error(e); process.exit(1); });
