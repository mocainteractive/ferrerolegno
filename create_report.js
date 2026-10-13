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
const FONT = 'Figtree';

function fmt(n) { return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function pct(a, b) { return ((a - b) / b * 100); }
function pctStr(a, b) { const v = pct(a, b); return (v > 0 ? '+' : '') + v.toFixed(1) + '%'; }

function addMocaLogo(slide) {
  slide.addShape(pres.shapes.OVAL, {
    x: 12.35, y: 6.95, w: 0.22, h: 0.22, fill: { color: RED }
  });
  slide.addText('moca', {
    x: 12.62, y: 6.9, w: 0.6, h: 0.3,
    fontSize: 12, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 7.0, w: 9, h: 0.3,
    fontSize: 8, fontFace: FONT, color: 'AAAAAA', isTextBox: true, margin: 0
  });
  addMocaLogo(slide);
}

function addLabel(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 0.35, w: 10, h: 0.3,
    fontSize: 10, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
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
// SLIDE 2 - SECTION: ANDAMENTO E RISULTATI
// =============================================
pres.addSection({ title: 'Andamento e risultati' });
const s2 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s2.addText([
  { text: 'Andamento e risultati', options: { fontSize: 38, fontFace: FONT, color: WHITE, italic: true, breakLine: true } },
  { text: 'Settembre 2026.', options: { fontSize: 38, fontFace: FONT, color: '888888', italic: true } }
], { x: 1.0, y: 2.5, w: 11, h: 2.5, isTextBox: true, margin: 0 });

// =============================================
// SLIDE 3 - TRAFFICO AL SITO | PANORAMICA
// =============================================
const s3 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s3, 'TRAFFICO AL SITO | PANORAMICA');

s3.addText('Forte crescita delle sessioni a settembre: +59% rispetto al 2025, trainata dai canali paid', {
  x: 0.5, y: 0.7, w: 12, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const days = Array.from({length: 30}, (_, i) => String(i + 1));
const sess2026 = [584,702,633,656,564,521,623,620,941,935,1046,938,1024,938,986,1076,985,952,969,955,966,953,975,937,885,892,910,907,894,863];
const sess2025 = [604,627,520,556,506,474,380,556,579,583,539,535,501,433,533,592,573,599,559,448,410,595,522,610,539,536,477,411,602,540];

s3.addChart(pres.charts.LINE, [
  { name: 'Sessioni settembre 2026', labels: days, values: sess2026 },
  { name: 'Sessioni settembre 2025', labels: days, values: sess2025 }
], {
  x: 0.5, y: 1.7, w: 8.0, h: 3.8,
  showTitle: false,
  showLegend: true, legendPos: 't', legendFontSize: 9, legendFontFace: FONT, legendColor: GRAY,
  chartColors: [RED, 'CCCCCC'],
  lineSize: 2.5,
  lineSmooth: true,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 7, catAxisLabelFontFace: FONT,
  valAxisLabelColor: GRAY, valAxisLabelFontSize: 8, valAxisLabelFontFace: FONT,
  valGridLine: { color: 'EEEEEE', size: 0.5 },
  catGridLine: { style: 'none' },
  showValue: false,
  catAxisOrientation: 'minMax',
  valAxisOrientation: 'minMax'
});

const kpis = [
  { label: 'SESSIONI', val2026: '25.691', val2025: '16.115', delta: '+59,4%', color: GREEN },
  { label: 'UTENTI', val2026: '18.350', val2025: '10.814', delta: '+69,7%', color: GREEN },
  { label: 'VISUALIZZAZIONI', val2026: '84.993', val2025: '77.872', delta: '+9,1%', color: GREEN },
  { label: 'TASSO COINVOLGIMENTO', val2026: '86,0%', val2025: '86,1%', delta: '-0,1pp', color: GRAY }
];

kpis.forEach((kpi, i) => {
  const xBase = 9.0 + (i < 2 ? 0 : 2.15);
  const yBase = 1.7 + (i % 2 === 0 ? 0 : 1.9);

  s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: xBase, y: yBase, w: 2.0, h: 1.7, rectRadius: 0.1,
    fill: { color: LGRAY }
  });
  s3.addText(kpi.label, {
    x: xBase + 0.15, y: yBase + 0.15, w: 1.7, h: 0.25,
    fontSize: 7, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0
  });
  s3.addText(kpi.val2026, {
    x: xBase + 0.15, y: yBase + 0.4, w: 1.7, h: 0.45,
    fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s3.addText(kpi.delta, {
    x: xBase + 0.15, y: yBase + 0.9, w: 1.0, h: 0.25,
    fontSize: 11, fontFace: FONT, bold: true, color: kpi.color, isTextBox: true, margin: 0
  });
  s3.addText('vs ' + kpi.val2025, {
    x: xBase + 0.15, y: yBase + 1.15, w: 1.7, h: 0.25,
    fontSize: 8, fontFace: FONT, color: 'AAAAAA', isTextBox: true, margin: 0
  });
});

s3.addText('Nonostante un avvio lento nella prima settimana, le sessioni hanno raggiunto picchi sopra le 1.000 sessioni/giorno dalla seconda settimana in poi, con un trend stabile e superiore al 2025.', {
  x: 0.5, y: 5.8, w: 8.5, h: 0.7,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

addFooter(s3, 'Fonte dati: GA4 | Periodo analizzato: set 2026 YoY');

// =============================================
// SLIDE 4 - TRAFFICO | CANALI DI ACQUISIZIONE
// =============================================
const s4 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s4, 'TRAFFICO PAID | PANORAMICA');

s4.addText('Il paid guida il 70% del traffico totale, con Paid Search e Cross-network in fortissima crescita', {
  x: 0.5, y: 0.7, w: 7.5, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const channelRows = [
  ['Canale di traffico', '', 'Sessioni 2026', '% sul tot.', 'Sessioni 2025', 'Var. %'],
  ['Paid Search', '', '8.129', '31,6%', '3.876', '+109,7%'],
  ['Paid Social', '', '5.933', '23,1%', '4.676', '+26,9%'],
  ['Cross-network', '', '3.917', '15,2%', '19', '+20.516%'],
  ['Organic Search', '', '4.271', '16,6%', '4.448', '-4,0%'],
  ['Direct', '', '1.914', '7,5%', '2.009', '-4,7%'],
  ['Organic Social', '', '234', '0,9%', '410', '-42,9%'],
  ['Altro', '', '1.293', '5,0%', '677', '+91,0%'],
  ['Totale complessivo', '', '25.691', '100%', '16.115', '+59,4%']
];

s4.addTable(channelRows, {
  x: 0.5, y: 1.7, w: 7.5,
  fontSize: 10, fontFace: FONT, color: DARK,
  border: { type: 'none' },
  colW: [1.8, 0.05, 1.2, 0.9, 1.2, 0.9],
  autoPage: false,
  rowH: [0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.4],
  rows: channelRows.map((_, i) => {
    if (i === 0) return { fill: { color: DARK }, color: WHITE, bold: true, fontSize: 9 };
    if (i === channelRows.length - 1) return { fill: { color: LGRAY }, bold: true };
    return { fill: { color: i % 2 === 0 ? LGRAY : WHITE } };
  })
});

s4.addChart(pres.charts.BAR, [
  { name: '2026', labels: ['Paid Search', 'Paid Social', 'Cross-net.', 'Org. Search', 'Direct'], values: [8129, 5933, 3917, 4271, 1914] },
  { name: '2025', labels: ['Paid Search', 'Paid Social', 'Cross-net.', 'Org. Search', 'Direct'], values: [3876, 4676, 19, 4448, 2009] }
], {
  x: 8.3, y: 1.7, w: 4.5, h: 4.0,
  showTitle: false,
  barDir: 'bar',
  barGrouping: 'clustered',
  chartColors: [RED, 'CCCCCC'],
  showValue: false,
  showLegend: true, legendPos: 't', legendFontSize: 8, legendFontFace: FONT, legendColor: GRAY,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 8, catAxisLabelFontFace: FONT,
  valAxisLabelColor: GRAY, valAxisLabelFontSize: 7, valAxisLabelFontFace: FONT,
  valGridLine: { color: 'EEEEEE', size: 0.5 },
  catGridLine: { style: 'none' },
  catAxisOrientation: 'maxMin'
});

s4.addText([
  { text: 'Forte crescita generale delle sessioni da Paid Search (+110%). ', options: { bold: true } },
  { text: 'In generale il paid porta il 70% del traffico totale. Cross-network e Organic Social: menzione particolare al canale in cui confluiscono campagne Demand Gen, PMax, YouTube e Display, cresciuto enormemente rispetto allo scorso anno.' }
], {
  x: 0.5, y: 5.9, w: 12, h: 0.8,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

addFooter(s4, 'Fonte dati: GA4 | Periodo analizzato: set 2026 YoY');

// =============================================
// SLIDE 5 - TRAFFICO ORGANICO | PANORAMICA
// =============================================
const s5 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s5, 'TRAFFICO ORGANICO | PANORAMICA');

s5.addText('Il traffico organico continua il trend di calo YoY, con clic in diminuzione e impression in lieve crescita', {
  x: 0.5, y: 0.7, w: 12, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const gscMetrics = [
  { label: 'CLIC', val: '5.391', delta: '-18%', deltaColor: RED, sub: 'vs 6.593 set 2025' },
  { label: 'IMPRESSION', val: '123.803', delta: '+1,7%', deltaColor: GREEN, sub: 'vs 121.681 set 2025' },
  { label: 'CTR MEDIO', val: '4,35%', delta: '-1,07pp', deltaColor: RED, sub: 'vs 5,42% set 2025' },
  { label: 'POSIZIONE MEDIA', val: '11,9', delta: '+4 pos.', deltaColor: GREEN, sub: 'vs 15,9 set 2025' }
];

gscMetrics.forEach((m, i) => {
  const xBase = 0.5 + i * 3.1;

  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: xBase, y: 1.8, w: 2.8, h: 2.5, rectRadius: 0.1,
    fill: { color: LGRAY }
  });
  s5.addText(m.label, {
    x: xBase + 0.2, y: 1.95, w: 2.4, h: 0.3,
    fontSize: 9, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0
  });
  s5.addText(m.val, {
    x: xBase + 0.2, y: 2.35, w: 2.4, h: 0.7,
    fontSize: 36, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s5.addText(m.delta, {
    x: xBase + 0.2, y: 3.1, w: 2.4, h: 0.35,
    fontSize: 16, fontFace: FONT, bold: true, color: m.deltaColor, isTextBox: true, margin: 0
  });
  s5.addText(m.sub, {
    x: xBase + 0.2, y: 3.55, w: 2.4, h: 0.3,
    fontSize: 8, fontFace: FONT, color: 'AAAAAA', isTextBox: true, margin: 0
  });
});

s5.addText([
  { text: 'Rispetto all\'anno precedente registriamo un calo dei clic organici (-18%), nonostante le impression siano leggermente in crescita (+1,7%). ', options: {} },
  { text: 'CTR e posizione media migliorano rispettivamente a 4,35% e 11,9, ', options: { bold: true } },
  { text: 'segno che il posizionamento del sito migliora ma il volume complessivo di ricerche cala, in linea con il trend di settore. L\'erosione del CTR da AI Overview e il calo delle ricerche di brand continuano a incidere.' }
], {
  x: 0.5, y: 4.6, w: 12, h: 1.2,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

addFooter(s5, 'Fonte dati: GSC - IT | Periodo analizzato: set 2026 YoY');

// =============================================
// SLIDE 6 - FOCUS ADVERTISING
// =============================================
const s6 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s6, 'TRAFFICO AL SITO | FOCUS ADVERTISING');

s6.addText('Google consolida la crescita con CPC in calo, Meta mantiene i volumi con investimento stabile', {
  x: 0.5, y: 0.7, w: 12, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

// Google Card
s6.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.8, w: 5.9, h: 4.5, rectRadius: 0.1, fill: { color: LGRAY }
});
s6.addText('Google Ads', {
  x: 0.8, y: 2.0, w: 3, h: 0.4,
  fontSize: 16, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});
s6.addText('Google e il primo canale paid per traffico al sito', {
  x: 0.8, y: 2.4, w: 5, h: 0.3,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

const googleKpis = [
  ['CLIC', '44.231', '+465%'],
  ['IMPRESSION', '465.037', '+552%'],
  ['INVESTIMENTO', '5.193 €', '+102%'],
  ['CPC MEDIO', '0,12 €', '-64%']
];
googleKpis.forEach((kpi, i) => {
  const y = 3.0 + i * 0.75;
  s6.addText(kpi[0], { x: 0.8, y: y, w: 1.8, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s6.addText(kpi[1], { x: 0.8, y: y + 0.2, w: 2.0, h: 0.35, fontSize: 18, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s6.addText(kpi[2], { x: 2.9, y: y + 0.2, w: 1.0, h: 0.35, fontSize: 12, fontFace: FONT, bold: true, color: kpi[2].startsWith('-') ? RED : GREEN, isTextBox: true, margin: 0 });
});

// Meta Card
s6.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.8, y: 1.8, w: 5.9, h: 4.5, rectRadius: 0.1, fill: { color: LGRAY }
});
s6.addText('Meta Ads', {
  x: 7.1, y: 2.0, w: 3, h: 0.4,
  fontSize: 16, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});
s6.addText('Meta mantiene i volumi con CTR in miglioramento', {
  x: 7.1, y: 2.4, w: 5, h: 0.3,
  fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

const metaKpis = [
  ['LINK CLIC', '12.970', '-6,7%'],
  ['REACH', '663.579', '-8,0%'],
  ['INVESTIMENTO', '3.826 €', '+10,6%'],
  ['CPC MEDIO', '0,30 €', '+18,5%']
];
metaKpis.forEach((kpi, i) => {
  const y = 3.0 + i * 0.75;
  s6.addText(kpi[0], { x: 7.1, y: y, w: 1.8, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s6.addText(kpi[1], { x: 7.1, y: y + 0.2, w: 2.0, h: 0.35, fontSize: 18, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s6.addText(kpi[2], { x: 9.2, y: y + 0.2, w: 1.0, h: 0.35, fontSize: 12, fontFace: FONT, bold: true, color: kpi[2].startsWith('+') && !kpi[0].includes('CPC') ? GREEN : kpi[2].startsWith('-') && kpi[0].includes('CPC') ? GREEN : RED, isTextBox: true, margin: 0 });
});

addFooter(s6, 'Fonte dati: Google Ads, Meta Ads | Periodo analizzato: set 2026 YoY');

// =============================================
// SLIDE 7 - GOOGLE ADS | DETTAGLIO CAMPAGNE
// =============================================
const s7 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s7, 'GOOGLE ADS | DETTAGLIO CAMPAGNE');

s7.addText('Demand Gen e Search Generic sono i principali driver di crescita, con CPC contenuti', {
  x: 0.5, y: 0.7, w: 12, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const gadsHeader = [
  [
    { text: 'Campagna', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
    { text: 'Clic', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'Impression', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'CTR', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'Costo', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'CPC', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } }
  ]
];
const gadsData = [
  ['Search Generic (Consideration)', '17.384', '92.951', '18,7%', '2.827 €', '0,16 €'],
  ['DG Consideration Traffico', '19.737', '295.383', '6,7%', '1.251 €', '0,06 €'],
  ['DG Conversion Lead Gen', '2.288', '64.777', '3,5%', '636 €', '0,28 €'],
  ['Search Brand Product', '1.911', '5.122', '37,3%', '240 €', '0,13 €'],
  ['Search Brand Protection', '2.911', '6.804', '42,8%', '239 €', '0,08 €']
];
const gadsTotals = [
  [
    { text: 'Totale', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK } },
    { text: '44.231', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '465.037', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '9,5%', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '5.193 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '0,12 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } }
  ]
];

const gadsRows = [
  ...gadsHeader,
  ...gadsData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: { fontSize: 10, fontFace: FONT, color: DARK, align: j > 0 ? 'right' : 'left', fill: { color: i % 2 === 0 ? WHITE : LGRAY } }
  }))),
  ...gadsTotals
];

s7.addTable(gadsRows, {
  x: 0.5, y: 1.7, w: 12.3,
  border: { type: 'none' },
  colW: [3.5, 1.5, 1.5, 1.2, 1.5, 1.2],
  rowH: [0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.45],
  autoPage: false
});

s7.addText([
  { text: 'YoY: ', options: { bold: true, color: DARK } },
  { text: 'i clic complessivi crescono del +465% con un investimento che raddoppia (+102%), ma il CPC medio scende del 64% (da 0,33 a 0,12 €). Le campagne Demand Gen, assenti nel 2025, portano da sole oltre 22.000 clic a costi molto contenuti.' }
], {
  x: 0.5, y: 5.0, w: 12, h: 0.9,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

s7.addChart(pres.charts.BAR, [
  { name: 'Set 2026', labels: ['Search Generic', 'DG Traffico', 'DG Lead Gen', 'Brand Product', 'Brand Prot.'], values: [2827, 1251, 636, 240, 239] },
  { name: 'Set 2025', labels: ['Search Generic', 'DG Traffico', 'DG Lead Gen', 'Brand Product', 'Brand Prot.'], values: [2153, 0, 0, 210, 210] }
], {
  x: 0.5, y: 5.6, w: 12.3, h: 1.3,
  showTitle: false,
  barDir: 'bar', barGrouping: 'clustered',
  chartColors: [RED, 'CCCCCC'],
  showValue: false,
  showLegend: true, legendPos: 't', legendFontSize: 7, legendFontFace: FONT, legendColor: GRAY,
  catAxisLabelColor: GRAY, catAxisLabelFontSize: 8, catAxisLabelFontFace: FONT,
  valAxisHidden: true,
  valGridLine: { style: 'none' }, catGridLine: { style: 'none' }
});

addFooter(s7, 'Fonte dati: Google Ads | Periodo analizzato: set 2026 YoY');

// =============================================
// SLIDE 8 - META ADS | DETTAGLIO CAMPAGNE
// =============================================
const s8 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s8, 'META ADS | DETTAGLIO CAMPAGNE');

s8.addText('Blog e Lead Generation rivenditori concentrano la maggior parte dell\'investimento Meta', {
  x: 0.5, y: 0.7, w: 12, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const metaHeader = [
  [
    { text: 'Campagna', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, fill: { color: DARK } } },
    { text: 'Link clic', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'CTR', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'CPC', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } },
    { text: 'Spesa', options: { bold: true, color: WHITE, fontSize: 9, fontFace: FONT, align: 'right', fill: { color: DARK } } }
  ]
];
const metaData = [
  ['Lead Gen Rivenditori DEF', '2.308', '1,92%', '0,48 €', '1.108 €'],
  ['Blog Consideration', '7.059', '2,26%', '0,15 €', '1.061 €'],
  ['FEEL5 B2B', '575', '0,92%', '0,99 €', '572 €'],
  ['Iscrizioni Area Pro B2B', '491', '0,87%', '1,13 €', '556 €'],
  ['Promo PED', '429', '0,07%', '0,68 €', '292 €'],
  ['Store Locator B2C', '2.108', '2,01%', '0,11 €', '237 €']
];
const metaTotals = [
  [
    { text: 'Totale', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK } },
    { text: '12.970', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '1,01%', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '0,30 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } },
    { text: '3.826 €', options: { bold: true, fontSize: 10, fontFace: FONT, color: DARK, align: 'right' } }
  ]
];

const metaRows = [
  ...metaHeader,
  ...metaData.map((row, i) => row.map((cell, j) => ({
    text: cell,
    options: { fontSize: 10, fontFace: FONT, color: DARK, align: j > 0 ? 'right' : 'left', fill: { color: i % 2 === 0 ? WHITE : LGRAY } }
  }))),
  ...metaTotals
];

s8.addTable(metaRows, {
  x: 0.5, y: 1.7, w: 12.3,
  border: { type: 'none' },
  colW: [3.5, 1.5, 1.2, 1.2, 1.5],
  rowH: [0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.45],
  autoPage: false
});

s8.addText([
  { text: 'La campagna Blog (Consideration) ', options: { bold: true } },
  { text: 'porta il volume maggiore di clic (7.059) con un CPC molto efficiente di 0,15 €. La campagna Lead Gen Rivenditori genera traffico qualificato con un CTR del 1,92%. Store Locator si conferma efficiente con CPC di 0,11 €.' }
], {
  x: 0.5, y: 5.5, w: 12, h: 0.9,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

addFooter(s8, 'Fonte dati: Meta Ads | Periodo analizzato: set 2026');

// =============================================
// SLIDE 9 - TRAFFICO ORGANICO | DETTAGLIO
// =============================================
const s9 = pres.addSlide({ sectionTitle: 'Andamento e risultati' });
addLabel(s9, 'TRAFFICO ORGANICO | DETTAGLIO');

s9.addText('Clic e impression globali: la posizione migliora ma il CTR scende per effetto AI Overview', {
  x: 0.5, y: 0.7, w: 12, h: 0.8,
  fontSize: 20, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

// Global GSC stats
s9.addText('DATI GLOBALI (TUTTI I PAESI)', {
  x: 0.5, y: 1.7, w: 5, h: 0.3,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const gscGlobal = [
  { label: 'CLIC', v26: '6.738', v25: '7.771', delta: '-13,3%', color: RED },
  { label: 'IMPRESSION', v26: '220.993', v25: '197.214', delta: '+12,1%', color: GREEN },
  { label: 'CTR', v26: '3,05%', v25: '3,94%', delta: '-0,89pp', color: RED },
  { label: 'POSIZIONE', v26: '10,3', v25: '14,3', delta: '+4 pos.', color: GREEN }
];

gscGlobal.forEach((m, i) => {
  const x = 0.5 + i * 3.0;
  s9.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: 2.1, w: 2.7, h: 1.5, rectRadius: 0.08, fill: { color: LGRAY } });
  s9.addText(m.label, { x: x + 0.15, y: 2.2, w: 2.4, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s9.addText(m.v26, { x: x + 0.15, y: 2.5, w: 2.4, h: 0.45, fontSize: 24, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s9.addText(m.delta, { x: x + 0.15, y: 3.0, w: 1.5, h: 0.25, fontSize: 11, fontFace: FONT, bold: true, color: m.color, isTextBox: true, margin: 0 });
});

// Italy GSC stats
s9.addText('DATI ITALIA', {
  x: 0.5, y: 3.9, w: 5, h: 0.3,
  fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
});

const gscIT = [
  { label: 'CLIC', v26: '5.391', v25: '6.593', delta: '-18,2%', color: RED },
  { label: 'IMPRESSION', v26: '123.803', v25: '121.681', delta: '+1,7%', color: GREEN },
  { label: 'CTR', v26: '4,35%', v25: '5,42%', delta: '-1,07pp', color: RED },
  { label: 'POSIZIONE', v26: '11,9', v25: '15,9', delta: '+4 pos.', color: GREEN }
];

gscIT.forEach((m, i) => {
  const x = 0.5 + i * 3.0;
  s9.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: 4.3, w: 2.7, h: 1.5, rectRadius: 0.08, fill: { color: LGRAY } });
  s9.addText(m.label, { x: x + 0.15, y: 4.4, w: 2.4, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s9.addText(m.v26, { x: x + 0.15, y: 4.7, w: 2.4, h: 0.45, fontSize: 24, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0 });
  s9.addText(m.delta, { x: x + 0.15, y: 5.2, w: 1.5, h: 0.25, fontSize: 11, fontFace: FONT, bold: true, color: m.color, isTextBox: true, margin: 0 });
});

s9.addText('L\'Italia pesa per l\'80% dei clic organici. Nonostante il miglioramento della posizione media (da 15,9 a 11,9), i clic calano per effetto combinato della contrazione della domanda di settore, dell\'erosione del CTR da AI Overview e del calo delle ricerche di brand.', {
  x: 0.5, y: 6.0, w: 12, h: 0.7,
  fontSize: 11, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
});

addFooter(s9, 'Fonte dati: GSC | Periodo analizzato: set 2026 YoY');

// =============================================
// SLIDE 10 - SECTION: IN CHIUSURA
// =============================================
pres.addSection({ title: 'In chiusura' });
const s10 = pres.addSlide({ sectionTitle: 'In chiusura' });
s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: DARK } });
s10.addText('In chiusura.', {
  x: 1.0, y: 2.8, w: 11.3, h: 2,
  fontSize: 38, fontFace: FONT, color: WHITE, italic: true, isTextBox: true, margin: 0
});

// =============================================
// SLIDE 11 - RIEPILOGO E PROSSIMI PASSI
// =============================================
const s11 = pres.addSlide({ sectionTitle: 'In chiusura' });
addLabel(s11, 'RIEPILOGO | SETTEMBRE 2026');

s11.addText('Punti chiave e considerazioni', {
  x: 0.5, y: 0.7, w: 12, h: 0.7,
  fontSize: 22, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const points = [
  {
    title: 'CRESCITA TRAFFICO PAID',
    text: 'Le sessioni totali crescono del 59% YoY, trainate dall\'investimento paid. I canali a pagamento portano il 70% del traffico complessivo. In particolare, Search Generic (+353%) e le nuove campagne Demand Gen guidano la crescita con CPC molto efficiente (0,06-0,16 €).'
  },
  {
    title: 'CALO ORGANICO',
    text: 'Il traffico organico continua il trend di calo (-18% clic in Italia). Le cause principali restano la contrazione della domanda di settore, l\'erosione del CTR da AI Overview e il calo delle ricerche di brand. La posizione media migliora (da 15,9 a 11,9), segno che non si tratta di un problema di posizionamento.'
  },
  {
    title: 'META ADS',
    text: 'La campagna Blog si conferma il miglior canale per traffico qualificato a basso costo (CPC 0,15 €). La Lead Gen Rivenditori porta conversioni con un CTR del 1,92%. L\'investimento totale su Meta cresce del 10,6% a 3.826 €.'
  },
  {
    title: 'PROSSIMI PASSI',
    text: 'Definire i filoni da veicolare nelle campagne DG Traffico. Monitorare l\'impatto di AI Overview sui contenuti del magazine. Valutare l\'allocazione del budget sui canali paid in vista della stagione autunnale.'
  }
];

points.forEach((p, i) => {
  const y = 1.6 + i * 1.3;
  s11.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: y, w: 12.3, h: 1.15, rectRadius: 0.08, fill: { color: i === points.length - 1 ? 'FFF0EF' : LGRAY }
  });

  s11.addText(p.title, {
    x: 0.8, y: y + 0.1, w: 3, h: 0.25,
    fontSize: 9, fontFace: FONT, bold: true, color: RED, isTextBox: true, margin: 0
  });
  s11.addText(p.text, {
    x: 0.8, y: y + 0.35, w: 11.7, h: 0.7,
    fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
  });
});

addFooter(s11, '');

// =============================================
// SLIDE 12 - CONTATTI
// =============================================
pres.addSection({ title: 'Contatti' });
const s12 = pres.addSlide({ sectionTitle: 'Contatti' });

s12.addShape(pres.shapes.OVAL, {
  x: 4.5, y: 0.8, w: 0.6, h: 0.6, fill: { color: RED }
});
s12.addText('moca', {
  x: 5.2, y: 0.8, w: 3, h: 0.6,
  fontSize: 36, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
});

const team = [
  { name: 'Federico Rugger', role: 'Project Account', email: 'federico.rugger@mocainteractive.com' },
  { name: 'Valeria Pompili', role: 'Digital Advertising Specialist', email: 'valeria.pompili@mocainteractive.com' },
  { name: 'Daniele Bertini', role: 'Digital Analytics Specialist', email: 'daniele.bertini@mocainteractive.com' },
  { name: 'Alessio Trombone', role: 'Digital Advertising Specialist', email: 'alessio.trombone@mocainteractive.com' },
  { name: 'Tommaso Minello', role: 'SEO Specialist', email: 'tommaso.minello@mocainteractive.com' }
];

team.forEach((t, i) => {
  const col = i < 3 ? 0 : 1;
  const row = i < 3 ? i : i - 3;
  const x = 0.8 + col * 4.5;
  const y = 2.2 + row * 1.4;

  s12.addText(t.name, {
    x: x, y: y, w: 4, h: 0.3,
    fontSize: 14, fontFace: FONT, bold: true, color: DARK, isTextBox: true, margin: 0
  });
  s12.addText(t.role, {
    x: x, y: y + 0.3, w: 4, h: 0.25,
    fontSize: 10, fontFace: FONT, color: GRAY, isTextBox: true, margin: 0
  });
  s12.addText(t.email, {
    x: x, y: y + 0.55, w: 4, h: 0.25,
    fontSize: 10, fontFace: FONT, color: RED, isTextBox: true, margin: 0
  });
});

s12.addText('MOCA Interactive S.r.l. Societa Benefit | Via Cernaia, 2 - 35141 Padova | P.IVA 04230520287', {
  x: 0.5, y: 6.8, w: 12, h: 0.3,
  fontSize: 7, fontFace: FONT, color: 'AAAAAA', align: 'center', isTextBox: true, margin: 0
});

addMocaLogo(s12);

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
