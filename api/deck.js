import PptxGenJS from 'pptxgenjs';

// Colours matching Playbk Labs brand
const NAVY   = '141f2e';
const BLUE   = '1a3a5c';
const ACCENT = '4a90c2';
const GREEN  = '1D9E75';
const RED    = 'c0392b';
const AMBER  = 'EF9F27';
const WHITE  = 'FFFFFF';
const OFFWHITE = 'F4F5F7';
const MUTED  = '8A99A8';

function healthColour(score) {
  if (score >= 70) return GREEN;
  if (score >= 40) return AMBER;
  return RED;
}

function forecastColour(fc) {
  const map = { Commit: GREEN, Upside: ACCENT, Pipeline: AMBER, Omit: RED };
  return map[fc] || MUTED;
}

function addSlide(pptx, title, bgColor = OFFWHITE) {
  const slide = pptx.addSlide();
  // Background
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: bgColor } });
  // Left accent bar
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.06, h: '100%', fill: { color: NAVY } });
  // Title
  slide.addText(title, { x: 0.35, y: 0.22, w: 9, h: 0.45, fontSize: 22, bold: true, color: NAVY, fontFace: 'Calibri' });
  // Thin rule under title
  slide.addShape(pptx.ShapeType.rect, { x: 0.35, y: 0.7, w: 9.3, h: 0.02, fill: { color: ACCENT } });
  return slide;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  let body = '';
  for await (const chunk of req) body += chunk;
  const { opp } = JSON.parse(body);
  const d = opp.data || {};
  const snap = d.snapshot || {};
  const medd = d.meddpicc || {};
  const answers = medd.meddAnswers || {};
  const stk = d.stakeholders || [];
  const accountpov = d.accountpov || {};
  const opppov = d.opppov || {};
  const conviction = d.conviction || {};
  const map = d.map || {};
  const mapRows = (map.rows || []).slice(0, 6);

  // Score helpers
  const totalQ = 28;
  const yesCount = Object.values(answers).filter(v => v === 'yes').length;
  const healthScore = opp.healthScore || Math.round((yesCount / totalQ) * 100);
  const hColor = healthColour(healthScore);

  const convStatus = [conviction.deal, conviction.win, conviction.play].every(v => v === 'yes') ? 'Convicted'
    : [conviction.deal, conviction.win, conviction.play].some(v => v === 'no') ? 'Walk Away?'
    : 'Uncertain';
  const convColor = convStatus === 'Convicted' ? GREEN : convStatus === 'Walk Away?' ? RED : AMBER;

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Playbk Labs';

  // ── SLIDE 1: Cover ────────────────────────────────────────────────────────
  {
    const slide = pptx.addSlide();
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: NAVY } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.5, h: '100%', fill: { color: BLUE } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 3.2, w: 9.2, h: 0.04, fill: { color: ACCENT } });

    slide.addText('DEAL REVIEW', { x: 0.8, y: 1.0, w: 8.5, h: 0.5, fontSize: 13, color: ACCENT, bold: true, charSpacing: 4, fontFace: 'Calibri' });
    slide.addText(snap.firmName || opp.name || 'Opportunity', { x: 0.8, y: 1.6, w: 8.5, h: 1.1, fontSize: 40, bold: true, color: WHITE, fontFace: 'Calibri' });
    slide.addText(snap.firmType || '', { x: 0.8, y: 2.75, w: 8.5, h: 0.4, fontSize: 16, color: MUTED, fontFace: 'Calibri' });

    // Stats row
    const stats = [
      { label: 'TARGET ARR', value: snap.targetArr || '—' },
      { label: 'DEAL HEALTH', value: String(healthScore) },
      { label: 'CLOSE DATE', value: snap.renewalQ || '—' },
      { label: 'CONVICTION', value: convStatus },
    ];
    stats.forEach((s, i) => {
      const x = 0.8 + i * 2.4;
      slide.addShape(pptx.ShapeType.rect, { x, y: 3.6, w: 2.1, h: 1.1, fill: { color: BLUE }, line: { color: ACCENT, width: 0.5 }, rounding: true });
      slide.addText(s.value, { x, y: 3.7, w: 2.1, h: 0.55, align: 'center', fontSize: 22, bold: true, color: i === 1 ? hColor : i === 3 ? convColor : WHITE, fontFace: 'Calibri' });
      slide.addText(s.label, { x, y: 4.25, w: 2.1, h: 0.3, align: 'center', fontSize: 9, color: MUTED, charSpacing: 2, fontFace: 'Calibri' });
    });

    slide.addText('Prepared by Playbk Labs', { x: 0.8, y: 5.1, w: 8, h: 0.3, fontSize: 10, color: MUTED, fontFace: 'Calibri' });
  }

  // ── SLIDE 2: Situation Brief ──────────────────────────────────────────────
  {
    const slide = addSlide(pptx, 'Situation Brief');
    const brief = snap.notes || 'No situation brief recorded.';
    const journey = snap.journey || '—';
    const health = snap.healthStatus || '—';

    slide.addShape(pptx.ShapeType.rect, { x: 0.35, y: 0.85, w: 6.4, h: 3.8, fill: { color: WHITE }, line: { color: OFFWHITE, width: 1 }, rounding: true });
    slide.addText(brief, { x: 0.55, y: 1.0, w: 6.0, h: 3.5, fontSize: 14, color: NAVY, fontFace: 'Calibri', valign: 'top', wrap: true });

    // Right panel stats
    const pills = [
      { label: 'Journey Stage', val: journey },
      { label: 'Account Health', val: health },
      { label: 'AUM', val: snap.aum || '—' },
      { label: 'Current ARR', val: snap.currentArr || '$0' },
      { label: 'Target ARR', val: snap.targetArr || '—' },
    ];
    pills.forEach((p, i) => {
      const y = 0.9 + i * 0.75;
      slide.addText(p.label.toUpperCase(), { x: 7.0, y, w: 2.5, h: 0.22, fontSize: 8, color: MUTED, charSpacing: 1.5, fontFace: 'Calibri' });
      slide.addText(p.val, { x: 7.0, y: y + 0.22, w: 2.5, h: 0.38, fontSize: 15, bold: true, color: NAVY, fontFace: 'Calibri' });
    });
  }

  // ── SLIDE 3: Deal Conviction ──────────────────────────────────────────────
  {
    const slide = addSlide(pptx, 'Deal Conviction');
    const qs = [
      { q: 'Is there a deal?', ans: conviction.deal },
      { q: 'Can we win?',      ans: conviction.win },
      { q: 'Should we play?',  ans: conviction.play },
    ];
    qs.forEach((item, i) => {
      const y = 0.95 + i * 1.35;
      const col = item.ans === 'yes' ? GREEN : item.ans === 'no' ? RED : AMBER;
      const label = (item.ans || 'Not set').charAt(0).toUpperCase() + (item.ans || 'not set').slice(1);
      slide.addShape(pptx.ShapeType.rect, { x: 0.35, y, w: 9.3, h: 1.1, fill: { color: WHITE }, line: { color: OFFWHITE, width: 1 }, rounding: true });
      slide.addShape(pptx.ShapeType.rect, { x: 0.35, y, w: 0.18, h: 1.1, fill: { color: col }, rounding: true });
      slide.addText(item.q, { x: 0.75, y: y + 0.15, w: 6.5, h: 0.4, fontSize: 16, bold: true, color: NAVY, fontFace: 'Calibri' });
      slide.addText(label, { x: 7.8, y: y + 0.2, w: 1.6, h: 0.65, align: 'center', fontSize: 18, bold: true, color: col, fontFace: 'Calibri' });
    });
    if (conviction.note) {
      slide.addText('AE NOTE: ' + conviction.note, { x: 0.35, y: 5.0, w: 9.3, h: 0.5, fontSize: 11, color: MUTED, italic: true, fontFace: 'Calibri', wrap: true });
    }
  }

  // ── SLIDE 4: Stakeholders ─────────────────────────────────────────────────
  {
    const slide = addSlide(pptx, 'Stakeholder Map');
    const cols = 2;
    stk.slice(0, 6).forEach((s, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = 0.35 + col * 4.85;
      const y = 0.9 + row * 1.55;
      const sentColor = s.sentiment === 'green' ? GREEN : s.sentiment === 'amber' ? AMBER : s.sentiment === 'red' ? RED : MUTED;

      slide.addShape(pptx.ShapeType.rect, { x, y, w: 4.5, h: 1.35, fill: { color: WHITE }, line: { color: OFFWHITE, width: 1 }, rounding: true });
      slide.addShape(pptx.ShapeType.ellipse, { x: x + 0.12, y: y + 0.12, w: 0.22, h: 0.22, fill: { color: sentColor } });
      slide.addText(s.name || 'Unknown', { x: x + 0.45, y: y + 0.1, w: 3.8, h: 0.35, fontSize: 13, bold: true, color: NAVY, fontFace: 'Calibri' });
      slide.addText((s.role || '') + (s.title ? ' · ' + s.title : ''), { x: x + 0.45, y: y + 0.42, w: 3.8, h: 0.28, fontSize: 10, color: MUTED, fontFace: 'Calibri' });
      slide.addText(s.win || '—', { x: x + 0.15, y: y + 0.73, w: 4.1, h: 0.5, fontSize: 10, color: NAVY, fontFace: 'Calibri', wrap: true });
    });
  }

  // ── SLIDE 5: Qualification (MEDDPPICC) ────────────────────────────────────
  {
    const slide = addSlide(pptx, 'Qualification (MEDDPPICC)');
    const sections = [
      { label: 'Metrics',           id: 'metrics',    q: 4 },
      { label: 'Economic Buyer',    id: 'eb',         q: 4 },
      { label: 'Decision Criteria', id: 'dc',         q: 3 },
      { label: 'Decision Process',  id: 'dp',         q: 4 },
      { label: 'Paper Process',     id: 'pp',         q: 3, blocker: true },
      { label: 'Identify Pain',     id: 'pain',       q: 4 },
      { label: 'Champion',          id: 'champion',   q: 4 },
      { label: 'Competition',       id: 'competition',q: 2 },
    ];
    sections.forEach((s, i) => {
      const col = i % 4;
      const row = Math.floor(i / 4);
      const x = 0.35 + col * 2.45;
      const y = 0.9 + row * 2.1;
      let yes = 0;
      for (let q = 0; q < s.q; q++) {
        if (answers[s.id + '_' + q] === 'yes') yes++;
      }
      const pct = Math.round((yes / s.q) * 100);
      const color = pct >= 75 ? GREEN : pct >= 40 ? AMBER : RED;

      slide.addShape(pptx.ShapeType.rect, { x, y, w: 2.2, h: 1.8, fill: { color: WHITE }, line: { color: OFFWHITE, width: 1 }, rounding: true });
      slide.addText(s.label, { x, y: y + 0.1, w: 2.2, h: 0.4, align: 'center', fontSize: 10, bold: true, color: NAVY, fontFace: 'Calibri' });
      slide.addText(pct + '%', { x, y: y + 0.55, w: 2.2, h: 0.7, align: 'center', fontSize: 32, bold: true, color, fontFace: 'Calibri' });
      slide.addText(yes + '/' + s.q + ' answered', { x, y: y + 1.3, w: 2.2, h: 0.35, align: 'center', fontSize: 9, color: MUTED, fontFace: 'Calibri' });
      if (s.blocker) slide.addText('⚠ Blocker', { x, y: y + 1.55, w: 2.2, h: 0.2, align: 'center', fontSize: 8, color: RED, fontFace: 'Calibri' });
    });
  }

  // ── SLIDE 6: Account POV ──────────────────────────────────────────────────
  {
    const slide = addSlide(pptx, 'Account POV');
    const fields = [
      { label: 'Business Problem',         val: accountpov.problemStatement },
      { label: 'Why Change Now',           val: accountpov.whyChangeNow },
      { label: 'Cost of Doing Nothing',    val: accountpov.costOfDoingNothing },
      { label: 'Value Proposition',        val: accountpov.valueProposition },
    ];
    fields.forEach((f, i) => {
      const y = 0.9 + i * 1.2;
      slide.addText((f.label + '').toUpperCase(), { x: 0.35, y, w: 9.3, h: 0.25, fontSize: 8, color: ACCENT, charSpacing: 2, bold: true, fontFace: 'Calibri' });
      slide.addText(f.val || '—', { x: 0.35, y: y + 0.25, w: 9.3, h: 0.82, fontSize: 12, color: NAVY, fontFace: 'Calibri', wrap: true });
      if (i < fields.length - 1) {
        slide.addShape(pptx.ShapeType.rect, { x: 0.35, y: y + 1.12, w: 9.3, h: 0.01, fill: { color: OFFWHITE } });
      }
    });
  }

  // ── SLIDE 7: Mutual Action Plan ───────────────────────────────────────────
  {
    const slide = addSlide(pptx, 'Mutual Action Plan');
    const headers = ['Action', 'Owner', 'Due'];
    const colW = [5.8, 1.8, 1.5];
    const colX = [0.35, 6.15, 7.95];

    // Header row
    headers.forEach((h, i) => {
      slide.addShape(pptx.ShapeType.rect, { x: colX[i], y: 0.85, w: colW[i], h: 0.38, fill: { color: NAVY } });
      slide.addText(h, { x: colX[i] + 0.1, y: 0.88, w: colW[i] - 0.1, h: 0.32, fontSize: 10, bold: true, color: WHITE, fontFace: 'Calibri' });
    });

    mapRows.forEach((r, i) => {
      const y = 1.28 + i * 0.65;
      const bg = i % 2 === 0 ? WHITE : OFFWHITE;
      slide.addShape(pptx.ShapeType.rect, { x: 0.35, y, w: 9.3, h: 0.6, fill: { color: bg } });
      const statusColor = r.status === 'Done' ? GREEN : r.status === 'At risk' ? RED : NAVY;
      slide.addText(r.action || '—', { x: colX[0] + 0.1, y: y + 0.1, w: colW[0] - 0.15, h: 0.45, fontSize: 11, color: statusColor, fontFace: 'Calibri', wrap: true });
      slide.addText(r.ownerAccount || r.ownerEleven || '—', { x: colX[1] + 0.05, y: y + 0.1, w: colW[1] - 0.05, h: 0.45, fontSize: 11, color: NAVY, fontFace: 'Calibri' });
      slide.addText(r.due || '—', { x: colX[2] + 0.05, y: y + 0.1, w: colW[2] - 0.05, h: 0.45, fontSize: 11, color: MUTED, fontFace: 'Calibri' });
    });
    if (mapRows.length === 0) {
      slide.addText('No MAP actions recorded yet.', { x: 0.35, y: 2.5, w: 9.3, h: 0.5, align: 'center', fontSize: 13, color: MUTED, fontFace: 'Calibri' });
    }
  }

  // ── SLIDE 8: Next Steps ───────────────────────────────────────────────────
  {
    const slide = pptx.addSlide();
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: NAVY } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.5, h: '100%', fill: { color: BLUE } });
    slide.addText('NEXT STEPS', { x: 0.8, y: 0.7, w: 9, h: 0.5, fontSize: 12, color: ACCENT, bold: true, charSpacing: 4, fontFace: 'Calibri' });

    const openActions = (map.rows || []).filter(r => r.status !== 'Done').slice(0, 4);
    if (openActions.length > 0) {
      openActions.forEach((r, i) => {
        const y = 1.4 + i * 1.1;
        slide.addShape(pptx.ShapeType.rect, { x: 0.8, y, w: 8.8, h: 0.9, fill: { color: BLUE }, rounding: true });
        slide.addShape(pptx.ShapeType.ellipse, { x: 0.9, y: y + 0.22, w: 0.45, h: 0.45, fill: { color: ACCENT } });
        slide.addText(String(i + 1), { x: 0.9, y: y + 0.22, w: 0.45, h: 0.45, align: 'center', fontSize: 13, bold: true, color: WHITE, fontFace: 'Calibri' });
        slide.addText(r.action || '—', { x: 1.5, y: y + 0.08, w: 5.8, h: 0.42, fontSize: 13, bold: true, color: WHITE, fontFace: 'Calibri' });
        slide.addText((r.ownerAccount || r.ownerEleven || '') + (r.due ? '  ·  Due: ' + r.due : ''), {
          x: 1.5, y: y + 0.5, w: 7.5, h: 0.3, fontSize: 10, color: MUTED, fontFace: 'Calibri'
        });
      });
    } else {
      slide.addText('Add actions in the Mutual Action Plan to populate this slide.', {
        x: 0.8, y: 2.5, w: 9, h: 0.5, align: 'center', fontSize: 14, color: MUTED, fontFace: 'Calibri'
      });
    }
    slide.addText('Generated by Playbk Labs  ·  playbklabs.ai', {
      x: 0.8, y: 5.2, w: 9, h: 0.3, fontSize: 9, color: MUTED, fontFace: 'Calibri'
    });
  }

  // ── Output ────────────────────────────────────────────────────────────────
  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  const filename = (opp.name || 'deal_review').replace(/\s+/g, '_') + '_Deal_Review.pptx';

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Length', buffer.length);
  res.status(200).send(buffer);
}
