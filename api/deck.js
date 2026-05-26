import PptxGenJS from 'pptxgenjs';

// ── Brand palette ─────────────────────────────────────────────────────────────
const C = {
  navy:    '0D1B2A',
  blue:    '1a3a5c',
  accent:  '4a90c2',
  green:   '1D9E75',
  red:     'C0392B',
  amber:   'EF9F27',
  white:   'FFFFFF',
  offwhite:'F7F8FA',
  light:   'EDF0F4',
  muted:   '8A99A8',
  text:    '1C2B3A',
  border:  'DDE3EA',
};

const FONT = 'Calibri';

function pct(yes, total) { return total ? Math.round((yes / total) * 100) : 0; }
function hColor(score) { return score >= 70 ? C.green : score >= 40 ? C.amber : C.red; }
function fcColor(fc) { return { Commit: C.green, Upside: C.accent, Pipeline: C.amber, Omit: C.red }[fc] || C.muted; }
function sentColor(s) { return { green: C.green, amber: C.amber, red: C.red }[s] || C.muted; }
function priorityColor(p) { return { High: C.red, Medium: C.amber, Low: C.green }[p] || C.muted; }

function slideHeader(slide, num, title) {
  // Section number badge
  slide.addShape('rect', { x: 0.4, y: 0.25, w: 0.45, h: 0.3, fill: { color: C.accent }, rectRadius: 0.04 });
  slide.addText(String(num).padStart(2, '0'), { x: 0.4, y: 0.25, w: 0.45, h: 0.3, align: 'center', fontSize: 10, bold: true, color: C.white, fontFace: FONT });
  slide.addText(title, { x: 1.0, y: 0.25, w: 8.5, h: 0.32, fontSize: 18, bold: true, color: C.text, fontFace: FONT });
}

function progressBar(slide, x, y, w, pctVal, color) {
  const h = 0.1;
  slide.addShape('rect', { x, y, w, h, fill: { color: C.light }, rectRadius: 0.05 });
  if (pctVal > 0) slide.addShape('rect', { x, y, w: w * pctVal / 100, h, fill: { color }, rectRadius: 0.05 });
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();
  let body = '';
  for await (const chunk of req) body += chunk;
  const { opp } = JSON.parse(body);
  const d    = opp.data || {};
  const snap = d.snapshot   || {};
  const medd = d.meddpicc   || {};
  const ans  = medd.meddAnswers || {};
  const stk  = d.stakeholders  || [];
  const apov = d.accountpov    || {};
  const opov = d.opppov        || {};
  const conv = d.conviction    || {};
  const map  = d.map           || {};
  const biz  = d.bizasks       || {};
  const calls= d.calls         || [];

  // Scores
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
  const totalQ = MEDD_SECTIONS.reduce((a, s) => a + s.q, 0);
  const totalY = MEDD_SECTIONS.reduce((a, s) => {
    for (let i = 0; i < s.q; i++) if (ans[s.id + '_' + i] === 'yes') a++;
    return a;
  }, 0);
  const healthScore = opp.healthScore || Math.round((totalY / totalQ) * 100);
  const hCol = hColor(healthScore);

  const convAnswers = [conv.deal, conv.win, conv.play];
  const convStatus = convAnswers.every(v => v === 'yes') ? 'Convicted'
    : convAnswers.some(v => v === 'no') ? 'Walk Away?' : 'Uncertain';

  // Latest coaching for slide 8
  const latestCall = [...calls].sort((a, b) => new Date(b.date) - new Date(a.date))[0];
  const coaching = latestCall?.coaching || {};
  const nba = coaching.actions || [];
  const stage = coaching.stage || '';

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Playbk Labs';

  // ── SLIDE 1: Cover ──────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    // Dark background
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.navy } });
    // Left stats panel
    s.addShape('rect', { x: 0, y: 0, w: 3.1, h: '100%', fill: { color: C.blue } });

    // Logo area
    s.addText('PLAYBK', { x: 0.25, y: 0.3, w: 2.6, h: 0.35, fontSize: 14, bold: true, color: C.white, charSpacing: 3, fontFace: FONT });
    s.addText('Deal Review', { x: 0.25, y: 0.65, w: 2.6, h: 0.28, fontSize: 11, color: C.accent, fontFace: FONT });

    // Stats
    const stats = [
      { label: 'FIRM TYPE',     val: snap.firmType   || '—' },
      { label: 'AUM',           val: snap.aum        || '—' },
      { label: 'CURRENT ARR',   val: snap.currentArr || '$0' },
      { label: 'TARGET ARR',    val: snap.targetArr  || '—' },
      { label: 'JOURNEY',       val: snap.journey    || '—' },
      { label: 'CLOSE / RENEW', val: snap.renewalQ   || '—' },
    ];
    stats.forEach((st, i) => {
      const y = 1.15 + i * 0.72;
      s.addShape('rect', { x: 0.15, y, w: 2.8, h: 0.02, fill: { color: 'FFFFFF', transparency: 80 } });
      s.addText(st.label, { x: 0.25, y: y + 0.07, w: 2.6, h: 0.2, fontSize: 8, color: C.muted, charSpacing: 1.5, fontFace: FONT });
      s.addText(st.val,   { x: 0.25, y: y + 0.28, w: 2.6, h: 0.28, fontSize: 13, bold: true, color: C.white, fontFace: FONT });
    });

    // Main title area
    s.addText(snap.firmName || opp.name || 'Opportunity', {
      x: 3.4, y: 1.6, w: 6.0, h: 1.3, fontSize: 38, bold: true, color: C.white, fontFace: FONT, wrap: true
    });
    s.addText(new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }), {
      x: 3.4, y: 2.95, w: 6.0, h: 0.35, fontSize: 13, color: C.muted, fontFace: FONT
    });

    // Health + Forecast badges
    s.addShape('ellipse', { x: 3.4, y: 3.5, w: 1.3, h: 1.3, fill: { color: C.blue }, line: { color: hCol, width: 3 } });
    s.addText(String(healthScore), { x: 3.4, y: 3.72, w: 1.3, h: 0.55, align: 'center', fontSize: 28, bold: true, color: hCol, fontFace: FONT });
    s.addText('HEALTH', { x: 3.4, y: 4.28, w: 1.3, h: 0.22, align: 'center', fontSize: 8, color: C.muted, charSpacing: 1, fontFace: FONT });

    const fc = snap.forecast || 'Pipeline';
    s.addShape('rect', { x: 5.1, y: 3.6, w: 1.5, h: 0.9, fill: { color: fcColor(fc), transparency: 80 }, rectRadius: 0.08 });
    s.addText(fc, { x: 5.1, y: 3.72, w: 1.5, h: 0.5, align: 'center', fontSize: 18, bold: true, color: fcColor(fc), fontFace: FONT });
    s.addText('FORECAST', { x: 5.1, y: 4.18, w: 1.5, h: 0.22, align: 'center', fontSize: 8, color: C.muted, charSpacing: 1, fontFace: FONT });

    s.addText('Confidential · Prepared with Playbk Labs', {
      x: 3.4, y: 5.25, w: 6.0, h: 0.28, fontSize: 9, color: C.muted, fontFace: FONT
    });
  }

  // ── SLIDE 2: Situation ──────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 1, 'SITUATION');
    s.addText(snap.firmName || opp.name || '', { x: 0.4, y: 0.62, w: 9.2, h: 0.32, fontSize: 22, bold: true, color: C.text, fontFace: FONT });

    // Situation brief
    s.addText(snap.notes || apov.problemStatement || 'No situation brief recorded.', {
      x: 0.4, y: 1.05, w: 5.8, h: 1.4, fontSize: 11, color: C.text, fontFace: FONT, wrap: true, valign: 'top'
    });

    // 4 why boxes
    const whys = [
      { label: 'WHY CHANGE?', val: apov.whyChangeNow || apov.problemStatement || '—' },
      { label: 'WHY NOW?',    val: apov.whyNow        || snap.notes || '—' },
      { label: 'WHY US?',     val: opov.winTheme      || apov.valueProposition || '—' },
      { label: 'WHAT SOLUTION?', val: opov.solution   || apov.valueProposition || '—' },
    ];
    whys.forEach((w, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.4 + col * 4.85;
      const y = 2.6 + row * 1.5;
      s.addShape('rect', { x, y, w: 4.6, h: 1.35, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.06 });
      s.addText(w.label, { x: x + 0.18, y: y + 0.1, w: 4.2, h: 0.22, fontSize: 8, bold: true, color: C.accent, charSpacing: 1.5, fontFace: FONT });
      s.addText(w.val,   { x: x + 0.18, y: y + 0.35, w: 4.2, h: 0.9, fontSize: 10, color: C.text, fontFace: FONT, wrap: true, valign: 'top' });
    });
  }

  // ── SLIDE 3: Stakeholder Map ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 2, 'STAKEHOLDER MAP');
    s.addText('Stakeholder Map', { x: 0.4, y: 0.62, w: 9.2, h: 0.32, fontSize: 22, bold: true, color: C.text, fontFace: FONT });

    // Stakeholder cards — 2 per row, up to 4
    stk.slice(0, 4).forEach((st, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.4 + col * 4.9;
      const y = 1.05 + row * 1.95;
      const sc = sentColor(st.sentiment);

      // Role badge
      const roleColors = { Champion: C.green, 'Economic buyer': C.amber, Coach: C.accent, Neutral: C.muted, Blocker: C.red };
      const rc = roleColors[st.role] || C.muted;
      s.addShape('rect', { x, y, w: 4.6, h: 1.75, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.06 });
      s.addShape('rect', { x, y, w: 1.2, h: 0.28, fill: { color: rc }, rectRadius: 0.04 });
      s.addText(st.role || '—', { x, y, w: 1.2, h: 0.28, align: 'center', fontSize: 9, bold: true, color: C.white, fontFace: FONT });

      // Sentiment dot
      s.addShape('ellipse', { x: x + 4.25, y: y + 0.05, w: 0.18, h: 0.18, fill: { color: sc } });
      const sentLabel = { green: 'Advocate', amber: 'Neutral', red: 'Risk' }[st.sentiment] || '—';
      s.addText(sentLabel, { x: x + 3.2, y: y + 0.04, w: 1.0, h: 0.22, align: 'right', fontSize: 8, color: sc, fontFace: FONT });

      s.addText(st.name  || 'Unknown',  { x: x + 0.15, y: y + 0.35, w: 4.2, h: 0.3,  fontSize: 14, bold: true, color: C.text, fontFace: FONT });
      s.addText(st.title || '',          { x: x + 0.15, y: y + 0.65, w: 4.2, h: 0.22, fontSize: 10, color: C.muted, fontFace: FONT });
      s.addText(st.win   || '—',         { x: x + 0.15, y: y + 0.92, w: 4.2, h: 0.72, fontSize: 10, color: C.text, fontFace: FONT, wrap: true, valign: 'top' });
    });

    // Coverage gap
    if (d.stakeholders?.length === 0 || stk[0]?.coverageGap || d.snapshot?.coverageGap) {
      const gap = '⚠  ' + (d.snapshot?.coverageGap || 'Review stakeholder coverage and access gaps.');
      s.addShape('rect', { x: 0.4, y: 5.0, w: 9.2, h: 0.52, fill: { color: 'FFF3CD' }, line: { color: C.amber, width: 0.5 }, rectRadius: 0.05 });
      s.addText(gap, { x: 0.6, y: 5.05, w: 8.8, h: 0.42, fontSize: 9, color: '7A5800', fontFace: FONT, wrap: true });
    }
  }

  // ── SLIDE 4: Opportunity POV ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 3, 'THE DEAL');
    s.addText('Opportunity POV', { x: 0.4, y: 0.62, w: 9.2, h: 0.32, fontSize: 22, bold: true, color: C.text, fontFace: FONT });

    const blocks = [
      { num: '1.', label: 'BUSINESS PROBLEM', val: opov.businessProblem || apov.problemStatement || '—' },
      { num: '2.', label: 'SUCCESS METRIC',   val: opov.successMetric   || opov.metrics || '—' },
      { num: '3.', label: 'THE ALTERNATIVE',  val: (opov.alternative    || '—') + (opov.competition ? '\n\nCompetition: ' + opov.competition : '') },
    ];
    blocks.forEach((b, i) => {
      const y = 1.1 + i * 1.45;
      s.addText(b.num + ' ' + b.label, { x: 0.4, y, w: 9.2, h: 0.25, fontSize: 9, bold: true, color: C.accent, charSpacing: 1.5, fontFace: FONT });
      s.addText(b.val, { x: 0.4, y: y + 0.28, w: 9.2, h: 1.0, fontSize: 11, color: C.text, fontFace: FONT, wrap: true, valign: 'top' });
      if (i < blocks.length - 1) s.addShape('rect', { x: 0.4, y: y + 1.35, w: 9.2, h: 0.015, fill: { color: C.border } });
    });

    if (opov.competition) {
      s.addText('Competition: ' + opov.competition, { x: 0.4, y: 5.25, w: 9.2, h: 0.3, fontSize: 9, color: C.muted, fontFace: FONT, italic: true });
    }
  }

  // ── SLIDE 5: Qualification ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 4, 'QUALIFICATION');
    s.addText('Qualification Scorecard', { x: 0.4, y: 0.62, w: 7.5, h: 0.32, fontSize: 22, bold: true, color: C.text, fontFace: FONT });

    // Health circle top-right
    s.addShape('ellipse', { x: 8.8, y: 0.35, w: 0.9, h: 0.9, fill: { color: C.offwhite }, line: { color: hCol, width: 3 } });
    s.addText(String(healthScore), { x: 8.8, y: 0.52, w: 0.9, h: 0.4, align: 'center', fontSize: 20, bold: true, color: hCol, fontFace: FONT });
    s.addText('HEALTH', { x: 8.75, y: 1.02, w: 1.0, h: 0.2, align: 'center', fontSize: 7, color: C.muted, charSpacing: 1, fontFace: FONT });

    MEDD_SECTIONS.forEach((sec, i) => {
      let yes = 0;
      for (let q = 0; q < sec.q; q++) if (ans[sec.id + '_' + q] === 'yes') yes++;
      const p = pct(yes, sec.q);
      const col = hColor(p);
      const note = medd.meddNotes?.[sec.id] || '';
      const isBlocker = sec.blocker;

      const y = 1.1 + i * 0.56;
      // Label
      const dotCol = isBlocker && p < 75 ? C.red : p >= 75 ? C.green : p >= 40 ? C.amber : C.red;
      s.addShape('ellipse', { x: 0.4, y: y + 0.08, w: 0.15, h: 0.15, fill: { color: dotCol } });
      s.addText(sec.label.toUpperCase(), { x: 0.65, y: y + 0.03, w: 2.5, h: 0.25, fontSize: 10, bold: true, color: C.text, fontFace: FONT });
      s.addText(yes + '/' + sec.q, { x: 3.2, y: y + 0.03, w: 0.6, h: 0.25, align: 'right', fontSize: 10, bold: true, color: col, fontFace: FONT });
      progressBar(s, 3.9, y + 0.1, 2.8, p, col);
      if (note) s.addText(note, { x: 6.85, y: y + 0.0, w: 2.8, h: 0.3, fontSize: 9, color: C.muted, fontFace: FONT, wrap: true });
    });
  }

  // ── SLIDE 6: Mutual Action Plan ─────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 5, 'ACTION PLAN');
    s.addText('Mutual Action Plan', { x: 0.4, y: 0.62, w: 9.2, h: 0.32, fontSize: 22, bold: true, color: C.text, fontFace: FONT });

    // Top ask callout
    const topAsk = (map.rows || []).find(r => r.status !== 'Done') || map.rows?.[0];
    if (topAsk) {
      s.addShape('rect', { x: 0.4, y: 1.05, w: 9.2, h: 0.55, fill: { color: 'FFF3CD' }, line: { color: C.amber, width: 0.5 }, rectRadius: 0.06 });
      s.addText('🚩  TOP ASK', { x: 0.6, y: 1.08, w: 1.5, h: 0.22, fontSize: 8, bold: true, color: C.red, fontFace: FONT });
      s.addText(topAsk.action || '—', { x: 2.1, y: 1.08, w: 7.3, h: 0.4, fontSize: 10, color: C.text, fontFace: FONT, wrap: true });
    }

    // Table headers
    const cols = [
      { label: 'DATE', x: 0.4,  w: 0.8  },
      { label: 'OPP OWNER', x: 1.25, w: 1.3  },
      { label: 'ACCOUNT OWNER', x: 2.6,  w: 1.7  },
      { label: 'ACTION / MILESTONE', x: 4.35, w: 4.2  },
      { label: 'STATUS', x: 8.6,  w: 1.0  },
    ];
    const tY = 1.72;
    s.addShape('rect', { x: 0.4, y: tY, w: 9.2, h: 0.35, fill: { color: C.navy } });
    cols.forEach(c => s.addText(c.label, { x: c.x + 0.08, y: tY + 0.05, w: c.w - 0.1, h: 0.25, fontSize: 8, bold: true, color: C.white, charSpacing: 1, fontFace: FONT }));

    (map.rows || []).slice(0, 5).forEach((r, i) => {
      const y = tY + 0.35 + i * 0.55;
      const bg = i % 2 === 0 ? C.white : C.offwhite;
      s.addShape('rect', { x: 0.4, y, w: 9.2, h: 0.52, fill: { color: bg } });
      const statusCol = r.status === 'Done' ? C.green : r.status === 'At risk' ? C.red : C.text;
      s.addText(r.due || '—',                            { x: cols[0].x + 0.08, y: y + 0.12, w: cols[0].w - 0.1, h: 0.3, fontSize: 9, color: C.text, fontFace: FONT });
      s.addText(r.ownerOpportunity || r.ownerEleven || '—', { x: cols[1].x + 0.08, y: y + 0.12, w: cols[1].w - 0.1, h: 0.3, fontSize: 9, color: C.text, fontFace: FONT });
      s.addText(r.ownerAccount || '—',                   { x: cols[2].x + 0.08, y: y + 0.12, w: cols[2].w - 0.1, h: 0.3, fontSize: 9, color: C.text, fontFace: FONT });
      s.addText(r.action || '—',                         { x: cols[3].x + 0.08, y: y + 0.08, w: cols[3].w - 0.1, h: 0.38, fontSize: 9, color: C.text, fontFace: FONT, wrap: true });
      s.addText(r.status || 'Open',                      { x: cols[4].x + 0.05, y: y + 0.12, w: cols[4].w - 0.05, h: 0.3, fontSize: 9, bold: true, color: statusCol, fontFace: FONT });
    });

    // Next touch note
    const nextRow = (map.rows || []).find(r => r.status === 'Open' && r.due);
    if (nextRow) {
      s.addText(`Next touch: ${nextRow.due} — ${nextRow.ownerOpportunity || nextRow.ownerEleven || ''} to ${(nextRow.action || '').toLowerCase().slice(0, 60)}.`, {
        x: 0.4, y: 5.22, w: 9.2, h: 0.3, fontSize: 9, color: C.muted, italic: true, fontFace: FONT
      });
    }
  }

  // ── SLIDE 7: Business Asks ──────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.white } });
    slideHeader(s, 6, 'INTERNAL ASKS');
    s.addText('Business Asks', { x: 0.4, y: 0.62, w: 9.2, h: 0.32, fontSize: 22, bold: true, color: C.text, fontFace: FONT });
    s.addText('What the account team needs from the business to close this deal.', {
      x: 0.4, y: 1.0, w: 9.2, h: 0.28, fontSize: 11, color: C.muted, fontFace: FONT
    });

    (biz.rows || []).slice(0, 4).forEach((r, i) => {
      const y = 1.42 + i * 1.05;
      const pc = priorityColor(r.priority);
      s.addShape('rect', { x: 0.4, y, w: 9.2, h: 0.9, fill: { color: C.offwhite }, line: { color: C.border, width: 0.5 }, rectRadius: 0.06 });
      // Team badge
      s.addShape('rect', { x: 0.55, y: y + 0.12, w: 0.9, h: 0.25, fill: { color: C.accent, transparency: 80 }, rectRadius: 0.04 });
      s.addText(r.team || '—', { x: 0.55, y: y + 0.12, w: 0.9, h: 0.25, align: 'center', fontSize: 8, bold: true, color: C.accent, fontFace: FONT });
      s.addText(r.ask || '—', { x: 1.55, y: y + 0.1, w: 6.5, h: 0.45, fontSize: 11, color: C.text, fontFace: FONT, wrap: true });
      // Priority + date
      s.addText(r.priority || 'Medium', { x: 1.55, y: y + 0.6, w: 1.2, h: 0.22, fontSize: 9, bold: true, color: pc, fontFace: FONT });
      s.addText(r.by ? '· By ' + r.by : '', { x: 2.8, y: y + 0.6, w: 6.8, h: 0.22, fontSize: 9, color: C.muted, fontFace: FONT });
    });
  }

  // ── SLIDE 8: Next Best Actions ──────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C.navy } });

    s.addText('07 · NEXT BEST ACTIONS', { x: 0.5, y: 0.3, w: 9.2, h: 0.3, fontSize: 11, color: C.accent, charSpacing: 2, bold: true, fontFace: FONT });
    s.addText('Next Best Actions', { x: 0.5, y: 0.65, w: 7, h: 0.5, fontSize: 26, bold: true, color: C.white, fontFace: FONT });
    if (stage) s.addText(stage.charAt(0).toUpperCase() + stage.slice(1) + ' stage', {
      x: 0.5, y: 1.18, w: 7, h: 0.3, fontSize: 12, color: C.muted, fontFace: FONT
    });

    if (nba.length === 0) {
      s.addText('Run coaching from the app to generate next best actions for this slide.', {
        x: 0.5, y: 2.8, w: 9.2, h: 0.5, align: 'center', fontSize: 13, color: C.muted, fontFace: FONT
      });
    } else {
      nba.slice(0, 3).forEach((a, i) => {
        const y = 1.6 + i * 1.3;
        s.addShape('rect', { x: 0.5, y, w: 9.2, h: 1.1, fill: { color: C.blue }, rectRadius: 0.08 });
        // Number circle
        s.addShape('ellipse', { x: 0.65, y: y + 0.27, w: 0.48, h: 0.48, fill: { color: C.accent } });
        s.addText(String(i + 1), { x: 0.65, y: y + 0.3, w: 0.48, h: 0.42, align: 'center', fontSize: 14, bold: true, color: C.white, fontFace: FONT });
        s.addText(a.title || a.action || '—', { x: 1.3, y: y + 0.1, w: 6.5, h: 0.38, fontSize: 13, bold: true, color: C.white, fontFace: FONT, wrap: true });
        s.addText(a.reason || a.rationale || '', { x: 1.3, y: y + 0.5, w: 6.5, h: 0.38, fontSize: 10, color: C.muted, fontFace: FONT, wrap: true });
        // Section tag
        const sectionMap = { snapshot:'Snapshot', stakeholders:'Stakeholders', accountpov:'Account', opppov:'Opportunity', meddpicc:'Qualification', map:'MAP', bizasks:'Biz Asks', calls:'Calls' };
        if (a.section) {
          const tag = sectionMap[a.section] || a.section;
          s.addShape('rect', { x: 8.1, y: y + 0.68, w: 1.4, h: 0.25, fill: { color: C.green, transparency: 70 }, rectRadius: 0.04 });
          s.addText(tag.toUpperCase(), { x: 8.1, y: y + 0.7, w: 1.4, h: 0.22, align: 'center', fontSize: 8, bold: true, color: C.green, fontFace: FONT });
        }
      });
    }

    s.addText('Generated by Playbk Labs · playbklabs.ai', {
      x: 0.5, y: 5.3, w: 9.2, h: 0.25, fontSize: 9, color: C.muted, fontFace: FONT
    });
  }

  // Output
  const buf = await pptx.write({ outputType: 'nodebuffer' });
  const filename = (opp.name || 'Deal_Review').replace(/\s+/g, '_') + '_Deal_Review.pptx';
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.status(200).send(buf);
}
