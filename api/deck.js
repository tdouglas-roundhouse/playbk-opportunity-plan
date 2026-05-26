import PptxGenJS from 'pptxgenjs';

const C = {
  navy:     '0D1B2A',
  blue:     '1a3a5c',
  accent:   '4a90c2',
  green:    '1D9E75',
  red:      'C0392B',
  amber:    'EF9F27',
  white:    'FFFFFF',
  offwhite: 'F7F8FA',
  light:    'EDF0F4',
  muted:    '8A99A8',
  text:     '1C2B3A',
  border:   'DDE3EA',
};
const F = 'Calibri';

function pct(yes, total) { return total ? Math.round((yes / total) * 100) : 0; }
function hColor(s) { return s >= 70 ? C.green : s >= 40 ? C.amber : C.red; }
function fcColor(f) { return { Commit: C.green, Upside: C.accent, Pipeline: C.amber, Omit: C.red }[f] || C.muted; }
function sentColor(s) { return { green: C.green, amber: C.amber, red: C.red }[s] || C.muted; }
function sentLabel(s) { return { green: 'Advocate', amber: 'Neutral', red: 'Risk' }[s] || '—'; }

const MEDD_SECTIONS = [
  { label: 'Metrics',           id: 'metrics',     q: 4 },
  { label: 'Economic Buyer',    id: 'eb',          q: 4 },
  { label: 'Decision Criteria', id: 'dc',          q: 3 },
  { label: 'Decision Process',  id: 'dp',          q: 4 },
  { label: 'Paper Process',     id: 'pp',          q: 3, blocker: true },
  { label: 'Identify Pain',     id: 'pain',        q: 4 },
  { label: 'Champion',          id: 'champion',    q: 4 },
  { label: 'Competition',       id: 'competition', q: 2 },
];

function slideHeader(s, num, section, title) {
  s.addShape('rect', { x: 0.4, y: 0.22, w: 0.52, h: 0.3, fill: { color: C.accent }, rectRadius: 0.04 });
  s.addText(String(num).padStart(2,'0'), { x: 0.4, y: 0.22, w: 0.52, h: 0.3, align: 'center', fontSize: 10, bold: true, color: C.white, fontFace: F });
  s.addText(section, { x: 1.05, y: 0.24, w: 4, h: 0.26, fontSize: 10, bold: true, color: C.muted, charSpacing: 2, fontFace: F });
  s.addText(title, { x: 0.4, y: 0.58, w: 9.2, h: 0.48, fontSize: 24, bold: true, color: C.text, fontFace: F });
}

function progressBar(s, x, y, w, pctVal, color) {
  s.addShape('rect', { x, y, w, h: 0.09, fill: { color: C.light }, rectRadius: 0.04 });
  if (pctVal > 0) s.addShape('rect', { x, y, w: w * pctVal / 100, h: 0.09, fill: { color }, rectRadius: 0.04 });
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();
  let body = '';
  for await (const chunk of req) body += chunk;
  const { opp } = JSON.parse(body);
  const d    = opp.data || {};
  const snap = d.snapshot    || {};
  const medd = d.meddpicc   || {};
  const ans  = medd.meddAnswers || {};
  const stk  = d.stakeholders  || [];
  const apov = d.accountpov    || {};
  const opov = d.opppov        || {};
  const conv = d.conviction    || {};
  const map  = d.map           || {};
  const biz  = d.bizasks       || {};
  const calls= d.calls         || [];

  const totalY = MEDD_SECTIONS.reduce((a, sec) => {
    for (let i = 0; i < sec.q; i++) if (ans[sec.id+'_'+i] === 'yes') a++;
    return a;
  }, 0);
  const totalQ = MEDD_SECTIONS.reduce((a, s) => a + s.q, 0);
  const healthScore = opp.healthScore || Math.round((totalY / totalQ) * 100);
  const hCol = hColor(healthScore);
  const fc = snap.forecast || 'Pipeline';

  const convAnswers = [conv.deal, conv.win, conv.play];
  const convStatus = convAnswers.every(v => v === 'yes') ? 'Convicted'
    : convAnswers.some(v => v === 'no') ? 'Walk Away?' : 'Uncertain';
  const convCol = convStatus === 'Convicted' ? C.green : convStatus === 'Walk Away?' ? C.red : C.amber;

  const champion = stk.find(s => s.role === 'Champion' || s.role === 'champion');
  const eb = stk.find(s => s.role === 'Economic buyer' || s.role === 'Economic Buyer');
  const blocker = stk.find(s => s.sentiment === 'red');
  const openActions = (map.rows || []).filter(r => r.status !== 'Done');
  const nextAction = openActions[0];
  const latestCall = [...calls].sort((a, b) => new Date(b.date) - new Date(a.date))[0];
  const nba = latestCall?.coaching?.actions || [];
  const stage = latestCall?.coaching?.stage || '';

  // Paper process score
  const ppYes = [0,1,2].filter(i => ans['pp_'+i] === 'yes').length;
  const ppPct = pct(ppYes, 3);

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Playbk Labs';

  // ── SLIDE 1: Cover ──────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.navy } });
    s.addShape('rect', { x: 0, y: 0, w: 3.2, h: '100%', fill: { color: C.blue } });

    s.addText('PLAYBK', { x: 0.3, y: 0.28, w: 2.6, h: 0.38, fontSize: 15, bold: true, color: C.white, charSpacing: 4, fontFace: F });
    s.addText('Deal Review', { x: 0.3, y: 0.68, w: 2.6, h: 0.28, fontSize: 11, color: C.accent, fontFace: F });

    const stats = [
      { label: 'FIRM TYPE',    val: snap.firmType   || '—' },
      { label: 'AUM',          val: snap.aum        || '—' },
      { label: 'CURRENT ARR',  val: snap.currentArr || '$0' },
      { label: 'TARGET ARR',   val: snap.targetArr  || '—' },
      { label: 'JOURNEY',      val: snap.journey    || '—' },
      { label: 'CLOSE DATE',   val: snap.renewalQ   || '—' },
    ];
    stats.forEach((st, i) => {
      const y = 1.2 + i * 0.72;
      s.addShape('rect', { x: 0.2, y, w: 2.8, h: 0.015, fill: { color: C.white, transparency: 75 } });
      s.addText(st.label, { x: 0.3, y: y + 0.07, w: 2.6, h: 0.2, fontSize: 8, color: C.muted, charSpacing: 1.5, fontFace: F });
      s.addText(st.val,   { x: 0.3, y: y + 0.28, w: 2.6, h: 0.3, fontSize: 14, bold: true, color: C.white, fontFace: F });
    });

    s.addText(snap.firmName || opp.name || 'Opportunity', { x: 3.5, y: 1.4, w: 6.0, h: 1.4, fontSize: 42, bold: true, color: C.white, fontFace: F, wrap: true });
    s.addText(new Date().toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' }), { x: 3.5, y: 2.9, w: 6.0, h: 0.35, fontSize: 13, color: C.muted, fontFace: F });

    s.addShape('ellipse', { x: 3.5, y: 3.4, w: 1.4, h: 1.4, fill: { color: C.navy }, line: { color: hCol, width: 3.5 } });
    s.addText(String(healthScore), { x: 3.5, y: 3.62, w: 1.4, h: 0.6, align: 'center', fontSize: 32, bold: true, color: hCol, fontFace: F });
    s.addText('HEALTH', { x: 3.5, y: 4.22, w: 1.4, h: 0.22, align: 'center', fontSize: 8, color: C.muted, charSpacing: 1, fontFace: F });

    s.addShape('rect', { x: 5.3, y: 3.5, w: 1.7, h: 1.0, fill: { color: fcColor(fc), transparency: 82 }, rectRadius: 0.08 });
    s.addText(fc, { x: 5.3, y: 3.65, w: 1.7, h: 0.55, align: 'center', fontSize: 20, bold: true, color: fcColor(fc), fontFace: F });
    s.addText('FORECAST', { x: 5.3, y: 4.18, w: 1.7, h: 0.22, align: 'center', fontSize: 8, color: C.muted, charSpacing: 1, fontFace: F });

    s.addText('Confidential · Prepared with Playbk Labs', { x: 3.5, y: 5.3, w: 6.0, h: 0.28, fontSize: 9, color: C.muted, fontFace: F });
  }

  // ── SLIDE 2: Deal Summary ───────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.navy } });

    s.addText('DEAL SUMMARY', { x: 0.5, y: 0.3, w: 9, h: 0.28, fontSize: 10, color: C.accent, charSpacing: 3, bold: true, fontFace: F });
    s.addText(snap.firmName || opp.name || '', { x: 0.5, y: 0.62, w: 9, h: 0.52, fontSize: 28, bold: true, color: C.white, fontFace: F });

    // Big metric cards row
    const metrics = [
      { label: 'Deal Health', val: String(healthScore), col: hCol },
      { label: 'Forecast',    val: fc,                  col: fcColor(fc) },
      { label: 'Conviction',  val: convStatus,          col: convCol },
      { label: 'Target ARR',  val: snap.targetArr||'—', col: C.white },
      { label: 'Close Date',  val: snap.renewalQ||'—',  col: C.white },
    ];
    metrics.forEach((m, i) => {
      const x = 0.5 + i * 2.02;
      s.addShape('rect', { x, y: 1.3, w: 1.85, h: 1.1, fill: { color: C.blue }, rectRadius: 0.08 });
      s.addText(m.val, { x, y: 1.38, w: 1.85, h: 0.58, align: 'center', fontSize: i < 3 ? 22 : 18, bold: true, color: m.col, fontFace: F });
      s.addText(m.label.toUpperCase(), { x, y: 1.96, w: 1.85, h: 0.25, align: 'center', fontSize: 8, color: C.muted, charSpacing: 1, fontFace: F });
    });

    // Key people row
    const people = [
      { label: 'CHAMPION', name: champion?.name || 'Not identified', sub: champion?.title || '', col: C.green },
      { label: 'ECONOMIC BUYER', name: eb?.name || 'Not identified', sub: eb?.title || '', col: C.amber },
      { label: 'TOP RISK', name: blocker?.name || (ppPct === 0 ? 'Paper Process' : 'No blocker identified'), sub: blocker?.title || (ppPct === 0 ? 'CCO not engaged' : ''), col: C.red },
    ];
    people.forEach((p, i) => {
      const x = 0.5 + i * 3.3;
      s.addShape('rect', { x, y: 2.62, w: 3.0, h: 1.0, fill: { color: C.blue }, rectRadius: 0.08 });
      s.addShape('rect', { x, y: 2.62, w: 0.18, h: 1.0, fill: { color: p.col }, rectRadius: 0.05 });
      s.addText(p.label, { x: x + 0.3, y: 2.68, w: 2.5, h: 0.22, fontSize: 8, color: p.col, charSpacing: 1.5, bold: true, fontFace: F });
      s.addText(p.name,  { x: x + 0.3, y: 2.9,  w: 2.5, h: 0.35, fontSize: 15, bold: true, color: C.white, fontFace: F });
      s.addText(p.sub,   { x: x + 0.3, y: 3.25, w: 2.5, h: 0.25, fontSize: 10, color: C.muted, fontFace: F });
    });

    // Situation brief
    s.addText('SITUATION', { x: 0.5, y: 3.78, w: 9, h: 0.22, fontSize: 8, color: C.accent, charSpacing: 2, bold: true, fontFace: F });
    s.addText(snap.notes || apov.problemStatement || '—', { x: 0.5, y: 4.02, w: 9, h: 0.6, fontSize: 11, color: 'AABBCC', fontFace: F, wrap: true });

    // Next action
    if (nextAction) {
      s.addShape('rect', { x: 0.5, y: 4.72, w: 9, h: 0.52, fill: { color: C.green, transparency: 85 }, rectRadius: 0.06 });
      s.addText('NEXT ACTION', { x: 0.65, y: 4.77, w: 1.5, h: 0.22, fontSize: 8, bold: true, color: C.green, fontFace: F });
      s.addText((nextAction.due ? nextAction.due + ' · ' : '') + nextAction.action, { x: 2.1, y: 4.77, w: 7.2, h: 0.38, fontSize: 11, color: C.white, fontFace: F, wrap: true });
    }
  }

  // ── SLIDE 3: Situation ──────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 1, 'SITUATION', snap.firmName || opp.name || 'Account');

    // Brief top left
    s.addText(snap.notes || apov.problemStatement || '—', { x: 0.4, y: 1.15, w: 5.7, h: 1.3, fontSize: 12, color: C.text, fontFace: F, wrap: true, valign: 'top' });

    // Right: quick stats panel
    const rstats = [
      { l:'Journey Stage', v: snap.journey    || '—' },
      { l:'Account Health', v: snap.healthStatus || '—' },
      { l:'AUM',           v: snap.aum        || '—' },
      { l:'Current ARR',   v: snap.currentArr || '$0' },
      { l:'Target ARR',    v: snap.targetArr  || '—' },
      { l:'Close Date',    v: snap.renewalQ   || '—' },
    ];
    rstats.forEach((r, i) => {
      const y = 1.15 + i * 0.72;
      s.addShape('rect', { x: 6.5, y, w: 3.1, h: 0.65, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.05 });
      s.addText(r.l.toUpperCase(), { x: 6.65, y: y + 0.07, w: 2.8, h: 0.2, fontSize: 8, color: C.muted, charSpacing: 1.2, fontFace: F });
      s.addText(r.v, { x: 6.65, y: y + 0.28, w: 2.8, h: 0.28, fontSize: 14, bold: true, color: C.text, fontFace: F });
    });

    // 4 why boxes filling the bottom
    const whys = [
      { label: 'WHY CHANGE?',     val: apov.whyChangeNow || apov.problemStatement || '—' },
      { label: 'WHY NOW?',        val: apov.whyNow || snap.notes || '—' },
      { label: 'WHY US?',         val: opov.winTheme || apov.valueProposition || '—' },
      { label: 'WHAT SOLUTION?',  val: opov.solution || apov.valueProposition || '—' },
    ];
    whys.forEach((w, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.4 + col * 3.0;
      const y = 2.62 + row * 1.6;
      s.addShape('rect', { x, y, w: 2.75, h: 1.45, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.06 });
      s.addText(w.label, { x: x+0.15, y: y+0.1, w: 2.45, h: 0.24, fontSize: 8, bold: true, color: C.accent, charSpacing: 1.5, fontFace: F });
      s.addText(w.val,   { x: x+0.15, y: y+0.36, w: 2.45, h: 0.98, fontSize: 10, color: C.text, fontFace: F, wrap: true, valign: 'top' });
    });
  }

  // ── SLIDE 4: Stakeholders ───────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 2, 'STAKEHOLDER MAP', 'Stakeholder Map');

    const roleColors = { Champion: C.green, 'Economic buyer': C.amber, 'Economic Buyer': C.amber, Coach: C.accent, Neutral: C.muted, Blocker: C.red };

    stk.slice(0, 4).forEach((st, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.4 + col * 4.9;
      const y = 1.1  + row * 2.3;
      const rc = roleColors[st.role] || C.muted;
      const sc = sentColor(st.sentiment);

      s.addShape('rect', { x, y, w: 4.6, h: 2.1, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.07 });
      // Role badge full-width top
      s.addShape('rect', { x, y, w: 1.6, h: 0.32, fill: { color: rc }, rectRadius: 0.05 });
      s.addText(st.role || '—', { x, y, w: 1.6, h: 0.32, align: 'center', fontSize: 10, bold: true, color: C.white, fontFace: F });
      // Sentiment
      s.addShape('ellipse', { x: x+4.3, y: y+0.07, w: 0.18, h: 0.18, fill: { color: sc } });
      s.addText(sentLabel(st.sentiment), { x: x+3.2, y: y+0.06, w: 1.08, h: 0.2, align: 'right', fontSize: 9, color: sc, fontFace: F });

      s.addText(st.name  || 'Unknown', { x: x+0.18, y: y+0.42, w: 4.2, h: 0.35, fontSize: 16, bold: true, color: C.text, fontFace: F });
      s.addText(st.title || '',         { x: x+0.18, y: y+0.78, w: 4.2, h: 0.25, fontSize: 10, color: C.muted, fontFace: F });
      // Personal win label
      s.addText('PERSONAL WIN', { x: x+0.18, y: y+1.1, w: 4.2, h: 0.2, fontSize: 8, bold: true, color: C.accent, charSpacing: 1.2, fontFace: F });
      s.addText(st.win || '—',  { x: x+0.18, y: y+1.3, w: 4.2, h: 0.65, fontSize: 10, color: C.text, fontFace: F, wrap: true, valign: 'top' });
    });

    // Coverage gap warning
    const gap = d.snapshot?.coverageGap || (stk.length > 0 && stk.find(s => s.sentiment === 'red') ? `⚠  ${stk.find(s => s.sentiment === 'red').name} has not engaged — potential compliance blocker.` : null);
    if (gap) {
      s.addShape('rect', { x: 0.4, y: 5.52, w: 9.2, h: 0.4, fill: { color: 'FFF3CD' }, line: { color: C.amber, width: 0.5 }, rectRadius: 0.05 });
      s.addText(gap, { x: 0.6, y: 5.56, w: 8.8, h: 0.3, fontSize: 9, color: '7A5800', fontFace: F });
    }
  }

  // ── SLIDE 5: Opportunity POV ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 3, 'THE DEAL', 'Opportunity POV');

    // Left column: Business Problem + Success Metric
    s.addShape('rect', { x: 0.4, y: 1.1, w: 5.5, h: 4.6, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.07 });

    s.addText('1. BUSINESS PROBLEM', { x: 0.6, y: 1.22, w: 5.1, h: 0.25, fontSize: 9, bold: true, color: C.accent, charSpacing: 1.5, fontFace: F });
    s.addText(opov.businessProblem || apov.problemStatement || '—', { x: 0.6, y: 1.5, w: 5.1, h: 1.5, fontSize: 12, color: C.text, fontFace: F, wrap: true, valign: 'top' });

    s.addShape('rect', { x: 0.6, y: 3.1, w: 5.1, h: 0.015, fill: { color: C.border } });

    s.addText('2. SUCCESS METRIC', { x: 0.6, y: 3.2, w: 5.1, h: 0.25, fontSize: 9, bold: true, color: C.accent, charSpacing: 1.5, fontFace: F });
    s.addText(opov.successMetric || '—', { x: 0.6, y: 3.48, w: 5.1, h: 1.2, fontSize: 12, color: C.text, fontFace: F, wrap: true, valign: 'top' });

    s.addShape('rect', { x: 0.6, y: 4.75, w: 5.1, h: 0.015, fill: { color: C.border } });

    s.addText('WIN THEME', { x: 0.6, y: 4.82, w: 5.1, h: 0.22, fontSize: 9, bold: true, color: C.green, charSpacing: 1.5, fontFace: F });
    s.addText(opov.winTheme || apov.valueProposition || '—', { x: 0.6, y: 5.06, w: 5.1, h: 0.5, fontSize: 11, color: C.text, fontFace: F, wrap: true, valign: 'top', italic: true });

    // Right column: Alternative + Competition
    s.addShape('rect', { x: 6.1, y: 1.1, w: 3.5, h: 2.2, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.07 });
    s.addText('3. THE ALTERNATIVE', { x: 6.28, y: 1.22, w: 3.15, h: 0.25, fontSize: 9, bold: true, color: C.accent, charSpacing: 1.5, fontFace: F });
    s.addText(opov.alternative || '—', { x: 6.28, y: 1.5, w: 3.15, h: 1.65, fontSize: 11, color: C.text, fontFace: F, wrap: true, valign: 'top' });

    s.addShape('rect', { x: 6.1, y: 3.5, w: 3.5, h: 2.2, fill: { color: C.navy }, rectRadius: 0.07 });
    s.addText('COMPETITION', { x: 6.28, y: 3.62, w: 3.15, h: 0.25, fontSize: 9, bold: true, color: C.accent, charSpacing: 1.5, fontFace: F });
    s.addText(opov.competition || 'Not documented', { x: 6.28, y: 3.9, w: 3.15, h: 1.6, fontSize: 11, color: C.white, fontFace: F, wrap: true, valign: 'top' });
  }

  // ── SLIDE 6: Qualification ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 4, 'QUALIFICATION', 'Qualification Scorecard');

    s.addShape('ellipse', { x: 8.7, y: 0.2, w: 1.0, h: 1.0, fill: { color: C.offwhite }, line: { color: hCol, width: 3 } });
    s.addText(String(healthScore), { x: 8.7, y: 0.38, w: 1.0, h: 0.45, align: 'center', fontSize: 22, bold: true, color: hCol, fontFace: F });
    s.addText('HEALTH', { x: 8.65, y: 1.05, w: 1.1, h: 0.2, align: 'center', fontSize: 7, color: C.muted, charSpacing: 1, fontFace: F });

    MEDD_SECTIONS.forEach((sec, i) => {
      let yes = 0;
      for (let q = 0; q < sec.q; q++) if (ans[sec.id+'_'+q] === 'yes') yes++;
      const p   = pct(yes, sec.q);
      const col = hColor(p);
      const note = medd.meddNotes?.[sec.id] || '';
      const y = 1.18 + i * 0.57;

      s.addShape('ellipse', { x: 0.4, y: y+0.07, w: 0.17, h: 0.17, fill: { color: col } });
      s.addText(sec.label.toUpperCase() + (sec.blocker ? ' ●' : ''), { x: 0.65, y: y+0.03, w: 2.5, h: 0.28, fontSize: 10, bold: true, color: C.text, fontFace: F });
      s.addText(yes+'/'+sec.q, { x: 3.2, y: y+0.03, w: 0.65, h: 0.28, align: 'right', fontSize: 11, bold: true, color: col, fontFace: F });
      progressBar(s, 3.95, y+0.1, 2.8, p, col);
      if (note) s.addText(note, { x: 6.9, y: y+0.0, w: 2.7, h: 0.35, fontSize: 9, color: C.muted, fontFace: F, wrap: true });
    });
  }

  // ── SLIDE 7: Mutual Action Plan ─────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 5, 'ACTION PLAN', 'Mutual Action Plan');

    // Top ask
    const topAsk = openActions[0] || (map.rows||[])[0];
    if (topAsk) {
      s.addShape('rect', { x: 0.4, y: 1.12, w: 9.2, h: 0.55, fill: { color: 'FFF3CD' }, line: { color: C.amber, width: 0.5 }, rectRadius: 0.06 });
      s.addText('🚩  TOP ASK', { x: 0.6, y: 1.16, w: 1.5, h: 0.22, fontSize: 8, bold: true, color: C.red, fontFace: F });
      s.addText(topAsk.action || '—', { x: 2.1, y: 1.15, w: 7.3, h: 0.42, fontSize: 10, color: C.text, fontFace: F, wrap: true });
    }

    const cols = [
      { l:'DATE',            x:0.4,  w:0.9  },
      { l:'OPP OWNER',       x:1.35, w:1.3  },
      { l:'ACCOUNT OWNER',   x:2.7,  w:1.7  },
      { l:'ACTION / MILESTONE', x:4.45, w:4.2 },
      { l:'STATUS',          x:8.7,  w:0.9  },
    ];
    const tY = 1.78;
    s.addShape('rect', { x: 0.4, y: tY, w: 9.2, h: 0.36, fill: { color: C.navy } });
    cols.forEach(c => s.addText(c.l, { x: c.x+0.1, y: tY+0.06, w: c.w-0.1, h: 0.25, fontSize: 8, bold: true, color: C.white, charSpacing: 1, fontFace: F }));

    (map.rows||[]).slice(0, 5).forEach((r, i) => {
      const y = tY + 0.36 + i * 0.57;
      s.addShape('rect', { x: 0.4, y, w: 9.2, h: 0.54, fill: { color: i % 2 === 0 ? C.white : C.offwhite } });
      const stCol = r.status === 'Done' ? C.green : r.status === 'At risk' ? C.red : C.text;
      s.addText(r.due||'—',                                    { x:cols[0].x+0.1, y:y+0.13, w:cols[0].w-0.1, h:0.3, fontSize:10, color:C.text, fontFace:F });
      s.addText(r.ownerEleven||r.ownerOpportunity||'—',        { x:cols[1].x+0.1, y:y+0.13, w:cols[1].w-0.1, h:0.3, fontSize:10, color:C.text, fontFace:F });
      s.addText(r.ownerAccount||'—',                           { x:cols[2].x+0.1, y:y+0.13, w:cols[2].w-0.1, h:0.3, fontSize:10, color:C.text, fontFace:F });
      s.addText(r.action||'—',                                 { x:cols[3].x+0.1, y:y+0.08, w:cols[3].w-0.1, h:0.42, fontSize:10, color:C.text, fontFace:F, wrap:true });
      s.addText(r.status||'Open',                              { x:cols[4].x+0.05, y:y+0.13, w:cols[4].w, h:0.3, fontSize:10, bold:true, color:stCol, fontFace:F });
    });

    const nr = openActions[0];
    if (nr) s.addText(`Next touch: ${nr.due||''} — ${nr.ownerEleven||nr.ownerOpportunity||''} to ${(nr.action||'').slice(0,55).toLowerCase()}.`, {
      x:0.4, y:5.25, w:9.2, h:0.28, fontSize:9, color:C.muted, italic:true, fontFace:F
    });
  }

  // ── SLIDE 8: Business Asks ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 6, 'INTERNAL ASKS', 'Business Asks');
    s.addText('What the account team needs from the business to close this deal.', { x:0.4, y:1.1, w:9.2, h:0.28, fontSize:11, color:C.muted, fontFace:F });

    const rows = (biz.rows||[]).slice(0, 4);
    const rowH = rows.length <= 3 ? 1.3 : 1.0;
    rows.forEach((r, i) => {
      const y = 1.5 + i * (rowH + 0.12);
      const pc = r.priority === 'High' ? C.red : r.priority === 'Medium' ? C.amber : C.green;
      s.addShape('rect', { x:0.4, y, w:9.2, h:rowH, fill:{ color:C.offwhite }, line:{ color:C.border, width:0.5 }, rectRadius:0.07 });
      // Left colour strip
      s.addShape('rect', { x:0.4, y, w:0.15, h:rowH, fill:{ color:pc }, rectRadius:0.04 });
      // Team badge
      s.addShape('rect', { x:0.65, y:y+0.15, w:1.1, h:0.3, fill:{ color:C.accent, transparency:82 }, rectRadius:0.04 });
      s.addText(r.team||'—', { x:0.65, y:y+0.15, w:1.1, h:0.3, align:'center', fontSize:9, bold:true, color:C.accent, fontFace:F });
      s.addText(r.ask||'—', { x:1.88, y:y+0.1, w:7.1, h:rowH-0.5, fontSize:12, color:C.text, fontFace:F, wrap:true, valign:'top' });
      s.addText((r.priority||'Medium') + (r.by ? '  ·  By ' + r.by : ''), { x:1.88, y:y+rowH-0.35, w:7.1, h:0.28, fontSize:10, bold:true, color:pc, fontFace:F });
    });
  }

  // ── SLIDE 9: Next Best Actions ──────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.navy } });

    s.addText('07 · NEXT BEST ACTIONS', { x:0.5, y:0.28, w:9, h:0.28, fontSize:10, color:C.accent, charSpacing:2, bold:true, fontFace:F });
    s.addText('Next Best Actions', { x:0.5, y:0.6, w:7, h:0.52, fontSize:28, bold:true, color:C.white, fontFace:F });
    if (stage) s.addText(stage.charAt(0).toUpperCase()+stage.slice(1)+' stage', { x:0.5, y:1.15, w:7, h:0.3, fontSize:12, color:C.muted, fontFace:F });

    if (nba.length === 0) {
      s.addText('Run coaching from the app to populate next best actions.', { x:0.5, y:2.8, w:9, h:0.5, align:'center', fontSize:13, color:C.muted, fontFace:F });
    } else {
      nba.slice(0, 3).forEach((a, i) => {
        const y = 1.58 + i * 1.35;
        s.addShape('rect', { x:0.5, y, w:9.2, h:1.18, fill:{ color:C.blue }, rectRadius:0.08 });
        s.addShape('ellipse', { x:0.65, y:y+0.32, w:0.5, h:0.5, fill:{ color:C.accent } });
        s.addText(String(i+1), { x:0.65, y:y+0.35, w:0.5, h:0.44, align:'center', fontSize:15, bold:true, color:C.white, fontFace:F });
        s.addText(a.title||a.action||'—', { x:1.3, y:y+0.1, w:6.5, h:0.42, fontSize:14, bold:true, color:C.white, fontFace:F, wrap:true });
        s.addText(a.reason||a.rationale||'', { x:1.3, y:y+0.55, w:6.5, h:0.48, fontSize:10, color:C.muted, fontFace:F, wrap:true });
        const sectionMap = { snapshot:'Snapshot', stakeholders:'Stakeholders', accountpov:'Account', opppov:'Opportunity', meddpicc:'Qualification', map:'MAP', bizasks:'Biz Asks', calls:'Calls' };
        if (a.section) {
          const tag = sectionMap[a.section] || a.section;
          s.addShape('rect', { x:8.2, y:y+0.73, w:1.3, h:0.26, fill:{ color:C.green, transparency:75 }, rectRadius:0.04 });
          s.addText(tag.toUpperCase(), { x:8.2, y:y+0.75, w:1.3, h:0.22, align:'center', fontSize:8, bold:true, color:C.green, fontFace:F });
        }
      });
    }
    s.addText('Generated by Playbk Labs · playbklabs.ai', { x:0.5, y:5.32, w:9, h:0.25, fontSize:9, color:C.muted, fontFace:F });
  }

  const buf = await pptx.write({ outputType: 'nodebuffer' });
  const fn = (opp.name||'Deal_Review').replace(/\s+/g,'_')+'_Deal_Review.pptx';
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition',`attachment; filename="${fn}"`);
  res.status(200).send(buf);
}
