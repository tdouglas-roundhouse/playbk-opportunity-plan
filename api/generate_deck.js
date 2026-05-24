const pptxgen = require('pptxgenjs');

// ── Palette ───────────────────────────────────────────────────────
const NAVY='0f1923',NAVY2='162030',TEAL='1D9E75',TEAL2='15755a',
      WHITE='FFFFFF',SLATE='8fa3b1',DIM='4a6070',AMBER='EF9F27',
      LIGHT_BG='F8F7F4';

const clean = s => {
  if (!s) return '';
  return String(s)
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&apos;/g, "'")
    .replace(/[\u2018\u2019]/g, "'")   // curly single quotes
    .replace(/[\u201C\u201D]/g, '"')   // curly double quotes
    .replace(/\u2014/g, '-')           // em dash
    .replace(/\u2013/g, '-')           // en dash
    .replace(/\u2026/g, '...')         // ellipsis
    .replace(/\u00A0/g, ' ')           // non-breaking space
    .replace(/[^\x00-\x7F\u00C0-\u024F]/g, '') // strip remaining non-latin
    .trim();
};
const trunc = (s,n) => { const c=clean(s); return c.length>n?c.slice(0,n-1)+'…':c; };
const makeShadow = () => ({type:'outer',blur:8,offset:2,angle:135,color:'000000',opacity:0.08});

function addSlideHeader(slide,label,pres) {
  slide.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.06,h:5.625,fill:{color:TEAL}});
  slide.addText(label.toUpperCase(),{x:0.3,y:0.2,w:9.4,h:0.3,fontSize:8,fontFace:'Calibri',color:SLATE,charSpacing:3,align:'left',margin:0});
}
function card(slide,pres,x,y,w,h,opts={}) {
  slide.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:opts.bg||'FFFFFF'},line:{color:opts.border||'E8E7E3',width:0.5},shadow:opts.shadow!==false?makeShadow():undefined});
}
function healthBadge(slide,pres,x,y,score) {
  const color=score>=70?TEAL:score>=40?AMBER:'a32d2d';
  const r=0.4;
  slide.addShape(pres.shapes.OVAL,{x:x-r,y:y-r,w:r*2,h:r*2,fill:{color:'F0F8F5'},line:{color,width:2.5}});
  slide.addText(String(score),{x:x-r,y:y-0.18,w:r*2,h:0.36,fontSize:18,fontFace:'Calibri',bold:true,color,align:'center',margin:0});
  slide.addText('HEALTH',{x:x-r,y:y+r-0.02,w:r*2,h:0.2,fontSize:6,fontFace:'Calibri',color:SLATE,align:'center',charSpacing:1.5,margin:0});
}
function forecastBadge(slide,pres,x,y,fc) {
  const colors={Commit:TEAL,Upside:'185fa5',Pipeline:AMBER,Omit:'a32d2d'};
  const bgs={Commit:'EAF3DE',Upside:'E6F1FB',Pipeline:'FAEEDA',Omit:'FCE8E8'};
  const c=colors[fc]||'a32d2d',bg=bgs[fc]||'FCE8E8';
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w:1.1,h:0.28,fill:{color:bg},line:{color:c,width:0.75},rectRadius:0.04});
  slide.addText(fc,{x,y,w:1.1,h:0.28,fontSize:9,fontFace:'Calibri',bold:true,color:c,align:'center',valign:'middle',margin:0});
}

function calcHealthScore(data) {
  const ma=(data.meddpicc&&data.meddpicc.meddAnswers)||{};
  const SECS=[
    {id:'economic_buyer',blocker:true,q:4},{id:'champion',blocker:true,q:4},
    {id:'paper_process',blocker:true,q:4},{id:'decision_process',blocker:true,q:4},
    {id:'data_integration',blocker:true,q:4},{id:'metrics',blocker:false,q:4},
    {id:'decision_criteria',blocker:false,q:4},{id:'competition',blocker:false,q:4}
  ];
  let mY=0,mT=0,bY=0,bT=0,comp=0;
  SECS.forEach(s=>{for(let i=0;i<s.q;i++){mT++;if(ma[s.id+'_'+i]==='yes')mY++;if(s.blocker){bT++;if(ma[s.id+'_'+i]==='yes')bY++;}}});
  if(data.snapshot.firmName)comp++;if((data.stakeholders||[]).length)comp++;
  if(data.accountpov&&data.accountpov.whyChange)comp++;if(data.opppov&&data.opppov.whitespace)comp++;
  if(data.map&&data.map.topAsk)comp++;
  let rec=0;if(data.map&&data.map.nextTouch)rec+=0.5;if(data.map&&data.map.lastEngagement)rec+=0.5;
  return Math.max(0,Math.min(100,Math.round((mT?mY/mT:0)*40+(bT?bY/bT:0)*30+(comp/5)*20+rec*10)));
}
function oppForecast(opp) {
  const ma=(opp.data.meddpicc&&opp.data.meddpicc.meddAnswers)||{};
  const SECS=[
    {id:'economic_buyer',blocker:true,q:4},{id:'champion',blocker:true,q:4},
    {id:'paper_process',blocker:true,q:4},{id:'decision_process',blocker:true,q:4},
    {id:'data_integration',blocker:true,q:4},{id:'metrics',blocker:false,q:4},
    {id:'decision_criteria',blocker:false,q:4},{id:'competition',blocker:false,q:4}
  ];
  let bY=0,bT=0,tY=0,tQ=0;
  SECS.forEach(s=>{for(let i=0;i<s.q;i++){tQ++;if(ma[s.id+'_'+i]==='yes')tY++;if(s.blocker){bT++;if(ma[s.id+'_'+i]==='yes')bY++;}}});
  const p=tQ?tY/tQ:0,bp=bT?bY/bT:0;
  if(bp>=0.75&&p>=0.70)return'Commit';if(bp>=0.55&&p>=0.50)return'Upside';
  if(p>=0.25||tY>=5)return'Pipeline';return'Omit';
}

function slideCover(pres,opp) {
  const s=opp.data.snapshot,slide=pres.addSlide();
  slide.background={color:NAVY};
  slide.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:3.2,h:5.625,fill:{color:NAVY2}});
  slide.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.08,h:5.625,fill:{color:TEAL}});
  slide.addText('PLAYBK',{x:0.2,y:0.3,w:2.7,h:0.3,fontSize:10,fontFace:'Calibri',bold:true,color:TEAL,charSpacing:5,margin:0});
  slide.addText('Deal Review',{x:0.2,y:0.58,w:2.7,h:0.25,fontSize:9,fontFace:'Calibri',color:SLATE,margin:0});
  const stats=[['Firm type',s.firmType||'—'],['AUM',s.aum||'—'],['Current ARR',s.currentArr||'$0'],['Target ARR',s.targetArr||'—'],['Journey',s.journey||'Land'],['Close / renew',s.renewalQ||'—']];
  stats.forEach(([lbl,val],i)=>{
    const y=1.2+i*0.62;
    slide.addText(lbl.toUpperCase(),{x:0.2,y,w:2.7,h:0.18,fontSize:7,fontFace:'Calibri',color:SLATE,charSpacing:2,margin:0});
    slide.addText(val,{x:0.2,y:y+0.2,w:2.7,h:0.3,fontSize:13,fontFace:'Calibri',bold:true,color:WHITE,margin:0});
  });
  slide.addText(trunc(clean(opp.name),40),{x:3.5,y:1.4,w:6.2,h:1.4,fontSize:38,fontFace:'Calibri',bold:true,color:WHITE,wrap:true,margin:0});
  const now=new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});
  slide.addText(now,{x:3.5,y:2.9,w:6.2,h:0.3,fontSize:11,fontFace:'Calibri',color:SLATE,margin:0});
  const hs=calcHealthScore(opp.data),fc=oppForecast(opp);
  healthBadge(slide,pres,7.0,4.1,hs);
  forecastBadge(slide,pres,8.0,3.95,fc);
  slide.addText('FORECAST',{x:8.0,y:4.28,w:1.1,h:0.15,fontSize:6,fontFace:'Calibri',color:SLATE,align:'center',charSpacing:1.5,margin:0});
  slide.addText('Confidential · Prepared with Playbk Labs',{x:3.5,y:5.25,w:6.2,h:0.2,fontSize:7,fontFace:'Calibri',color:DIM,margin:0});
}

function slideSituation(pres,opp) {
  const slide=pres.addSlide();slide.background={color:LIGHT_BG};
  addSlideHeader(slide,'01 · Situation',pres);
  slide.addText(trunc(clean(opp.name),50),{x:0.3,y:0.52,w:9.4,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
  const p=opp.data.accountpov||{},notes=clean(opp.data.snapshot.notes);
  if(notes){card(slide,pres,0.3,1.12,9.4,0.85);slide.addText(trunc(notes,300),{x:0.5,y:1.18,w:9.0,h:0.73,fontSize:11,fontFace:'Calibri',color:'2a2a28',wrap:true,margin:0});}
  [['Why change?',p.whyChange],['Why now?',p.whyNow],['Why us?',p.whyUs],['What solution?',p.whatSolution]].forEach(([lbl,val],i)=>{
    const x=0.3+(i%2)*4.75,y=2.12+Math.floor(i/2)*1.45;
    card(slide,pres,x,y,4.6,1.3);
    slide.addText(lbl.toUpperCase(),{x:x+0.15,y:y+0.1,w:4.3,h:0.2,fontSize:7,fontFace:'Calibri',color:TEAL,charSpacing:2,bold:true,margin:0});
    slide.addText(trunc(val||'Not yet filled',130),{x:x+0.15,y:y+0.32,w:4.3,h:0.88,fontSize:10,fontFace:'Calibri',color:'1a1a18',wrap:true,margin:0});
  });
}

function slideStakeholders(pres,opp) {
  const slide=pres.addSlide();slide.background={color:LIGHT_BG};
  addSlideHeader(slide,'02 · Stakeholder Map',pres);
  slide.addText('Stakeholder Map',{x:0.3,y:0.52,w:9.4,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
  const stks=(opp.data.stakeholders||[]).slice(0,6),cols=Math.min(stks.length,3),cw=cols>0?(9.4/cols)-0.1:3;
  const sentColors={Advocate:TEAL,Neutral:AMBER,Risk:'a32d2d'};
  const roleColors={'Economic buyer':'185fa5','Executive Sponsor':'6b46c1',Champion:TEAL,Coach:AMBER,Blocker:'a32d2d',Neutral:DIM};
  stks.forEach((st,i)=>{
    const col=i%3,row=Math.floor(i/3),x=0.3+col*(cw+0.1),y=1.18+row*1.95;
    const sc=sentColors[st.sentiment]||DIM,rc=roleColors[st.role]||DIM;
    card(slide,pres,x,y,cw,1.75);
    slide.addShape(pres.shapes.RECTANGLE,{x,y,w:cw,h:0.06,fill:{color:sc}});
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:x+0.12,y:y+0.12,w:1.3,h:0.22,fill:{color:rc},line:{color:rc},rectRadius:0.03});
    slide.addText(clean(st.role),{x:x+0.12,y:y+0.12,w:1.3,h:0.22,fontSize:7,fontFace:'Calibri',bold:true,color:WHITE,align:'center',valign:'middle',margin:0});
    slide.addText(clean(st.sentiment||'—'),{x:x+cw-1.0,y:y+0.13,w:0.88,h:0.2,fontSize:7,fontFace:'Calibri',bold:true,color:sc,align:'right',margin:0});
    slide.addText(trunc(st.name||'Unknown',28),{x:x+0.12,y:y+0.4,w:cw-0.24,h:0.32,fontSize:13,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
    slide.addText(trunc(st.title||'',35),{x:x+0.12,y:y+0.72,w:cw-0.24,h:0.22,fontSize:9,fontFace:'Calibri',color:DIM,margin:0});
    if(st.win)slide.addText(trunc(st.win,80),{x:x+0.12,y:y+0.98,w:cw-0.24,h:0.65,fontSize:8.5,fontFace:'Calibri',color:'2a2a28',wrap:true,italic:true,margin:0});
  });
  const gaps=clean(opp.data.snapshot.stakeholderGaps);
  if(gaps){
    const y=stks.length>3?5.0:3.15;
    slide.addShape(pres.shapes.RECTANGLE,{x:0.3,y,w:9.4,h:0.4,fill:{color:'FCE8E8'},line:{color:'F7C1C1'}});
    slide.addText('⚠  '+trunc(gaps,140),{x:0.45,y:y+0.05,w:9.1,h:0.3,fontSize:9,fontFace:'Calibri',color:'a32d2d',wrap:true,margin:0});
  }
}

function slideDeal(pres,opp) {
  const slide=pres.addSlide();slide.background={color:LIGHT_BG};
  addSlideHeader(slide,'03 · The Deal',pres);
  slide.addText('Opportunity POV',{x:0.3,y:0.52,w:9.4,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
  const o=opp.data.opppov||{};
  [['1. Business problem',o.businessProblem,TEAL],['2. Success metric',o.successMetric,'185fa5'],['3. The alternative',o.alternative,AMBER]].forEach(([lbl,val,accent],i)=>{
    const y=1.15+i*1.35;card(slide,pres,0.3,y,9.4,1.2);
    slide.addShape(pres.shapes.RECTANGLE,{x:0.3,y,w:0.07,h:1.2,fill:{color:accent}});
    slide.addText(lbl.toUpperCase(),{x:0.5,y:y+0.1,w:8.9,h:0.2,fontSize:7.5,fontFace:'Calibri',bold:true,color:accent,charSpacing:1.5,margin:0});
    slide.addText(trunc(val||'Not yet filled',220),{x:0.5,y:y+0.33,w:8.9,h:0.78,fontSize:10.5,fontFace:'Calibri',color:'1a1a18',wrap:true,margin:0});
  });
  if(o.incumbents)slide.addText('Competition: '+trunc(o.incumbents,120),{x:0.3,y:5.22,w:9.4,h:0.25,fontSize:8.5,fontFace:'Calibri',color:DIM,italic:true,margin:0});
}

function slideQualification(pres,opp,MEDD) {
  const slide=pres.addSlide();slide.background={color:LIGHT_BG};
  addSlideHeader(slide,'04 · Qualification',pres);
  const ma=(opp.data.meddpicc&&opp.data.meddpicc.meddAnswers)||{};
  const mn=(opp.data.meddpicc&&opp.data.meddpicc.meddNotes)||{};
  slide.addText('Qualification Scorecard',{x:0.3,y:0.52,w:7,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
  healthBadge(slide,pres,9.2,0.77,calcHealthScore(opp.data));
  MEDD.forEach((sec,i)=>{
    const col=i%2,row=Math.floor(i/2),x=0.3+col*4.8,y=1.18+row*1.05;
    const yes=sec.questions.filter((_,qi)=>ma[sec.id+'_'+qi]==='yes').length,tot=sec.questions.length;
    const pct=yes/tot,barColor=pct>=0.75?TEAL:pct>=0.5?AMBER:'a32d2d';
    card(slide,pres,x,y,4.6,0.9,{shadow:false});
    slide.addText(sec.title.toUpperCase()+(sec.blocker?' ●':''),{x:x+0.12,y:y+0.08,w:3.2,h:0.2,fontSize:7.5,fontFace:'Calibri',bold:true,color:sec.blocker&&pct<0.5?'a32d2d':NAVY,charSpacing:1,margin:0});
    slide.addText(`${yes}/${tot}`,{x:x+3.4,y:y+0.07,w:1.0,h:0.22,fontSize:11,fontFace:'Calibri',bold:true,color:barColor,align:'right',margin:0});
    slide.addShape(pres.shapes.RECTANGLE,{x:x+0.12,y:y+0.36,w:4.35,h:0.1,fill:{color:'E8E7E3'},line:{color:'E8E7E3'}});
    if(pct>0)slide.addShape(pres.shapes.RECTANGLE,{x:x+0.12,y:y+0.36,w:4.35*pct,h:0.1,fill:{color:barColor},line:{color:barColor}});
    const note=mn[sec.id];
    if(note)slide.addText(trunc(note,80),{x:x+0.12,y:y+0.54,w:4.35,h:0.28,fontSize:7.5,fontFace:'Calibri',color:DIM,wrap:true,italic:true,margin:0});
  });
}

function slideMAP(pres,opp) {
  const slide=pres.addSlide();slide.background={color:LIGHT_BG};
  addSlideHeader(slide,'05 · Action Plan',pres);
  slide.addText('Mutual Action Plan',{x:0.3,y:0.52,w:9.4,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
  const m=opp.data.map||{};
  if(m.topAsk){
    slide.addShape(pres.shapes.RECTANGLE,{x:0.3,y:1.12,w:9.4,h:0.62,fill:{color:'EAF3DE'},line:{color:'C0DD97',width:0.5}});
    slide.addShape(pres.shapes.RECTANGLE,{x:0.3,y:1.12,w:0.07,h:0.62,fill:{color:TEAL}});
    slide.addText('🚩  TOP ASK',{x:0.5,y:1.18,w:9.1,h:0.18,fontSize:7,fontFace:'Calibri',bold:true,color:TEAL2,charSpacing:2,margin:0});
    slide.addText(trunc(m.topAsk,180),{x:0.5,y:1.36,w:9.1,h:0.32,fontSize:10,fontFace:'Calibri',color:'1a1a18',wrap:true,margin:0});
  }
  const rows=(m.rows||[]).slice(0,7).map(r=>{
    const sc={Open:'E6F1FB','In progress':'FAEEDA',Done:'EAF3DE',Blocked:'FCE8E8'}[r.status]||'FFFFFF';
    return[
      {text:clean(r.date)||'—',options:{fontSize:9,fontFace:'Calibri',color:'1a1a18'}},
      {text:clean(r.ownerEleven)||'—',options:{fontSize:9,fontFace:'Calibri',color:'1a1a18'}},
      {text:clean(r.ownerAccount)||'—',options:{fontSize:9,fontFace:'Calibri',color:'1a1a18'}},
      {text:trunc(r.action,80)||'—',options:{fontSize:9,fontFace:'Calibri',color:'1a1a18'}},
      {text:clean(r.status)||'—',options:{fontSize:9,fontFace:'Calibri',color:'1a1a18',fill:{color:sc}}}
    ];
  });
  if(rows.length){
    const hdr=[['DATE'],['ELEVEN OWNER'],['ACCOUNT OWNER'],['ACTION / MILESTONE'],['STATUS']].map(([t])=>({text:t,options:{fill:{color:NAVY},color:'FFFFFF',bold:true,fontSize:8,fontFace:'Calibri'}}));
    slide.addTable([hdr,...rows],{x:0.3,y:1.88,w:9.4,colW:[0.9,1.4,1.5,4.3,1.3],border:{pt:0.5,color:'E8E7E3'}});
  }
  if(m.nextTouch)slide.addText('Next touch: '+clean(m.nextTouch),{x:0.3,y:5.25,w:9.4,h:0.22,fontSize:9,fontFace:'Calibri',color:DIM,italic:true,margin:0});
}

function slideBizAsks(pres,opp) {
  const slide=pres.addSlide();slide.background={color:LIGHT_BG};
  addSlideHeader(slide,'06 · Internal Asks',pres);
  slide.addText('Business Asks',{x:0.3,y:0.52,w:9.4,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:NAVY,margin:0});
  slide.addText('What the account team needs from the business to close this deal.',{x:0.3,y:1.05,w:9.4,h:0.28,fontSize:10,fontFace:'Calibri',color:DIM,margin:0});
  const rows=((opp.data.bizasks&&opp.data.bizasks.rows)||[]).slice(0,6);
  if(!rows.length){slide.addText('No business asks logged yet.',{x:0.3,y:2.0,w:9.4,h:0.4,fontSize:11,fontFace:'Calibri',color:SLATE,italic:true,margin:0});return;}
  const teamColors={Product:TEAL,Legal:AMBER,Executive:'185fa5',Marketing:'6b46c1',Finance:TEAL2,Engineering:DIM,Other:DIM};
  const priColors={High:'a32d2d',Medium:AMBER,Low:TEAL};
  rows.forEach((r,i)=>{
    const y=1.48+i*0.67,tc=teamColors[r.team]||DIM,sc=priColors[r.priority]||DIM;
    card(slide,pres,0.3,y,9.4,0.58,{shadow:false});
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.42,y:y+0.1,w:1.1,h:0.22,fill:{color:tc},line:{color:tc},rectRadius:0.03});
    slide.addText(clean(r.team),{x:0.42,y:y+0.1,w:1.1,h:0.22,fontSize:8,fontFace:'Calibri',bold:true,color:'FFFFFF',align:'center',valign:'middle',margin:0});
    slide.addText(trunc(r.ask,130),{x:1.65,y:y+0.07,w:6.3,h:0.44,fontSize:10,fontFace:'Calibri',color:'1a1a18',wrap:true,margin:0});
    slide.addText((r.priority||'')+(r.by?' · By '+clean(r.by):''),{x:8.05,y:y+0.1,w:1.5,h:0.38,fontSize:8,fontFace:'Calibri',color:sc,align:'right',bold:true,margin:0});
  });
}

function slideCoaching(pres,opp) {
  const slide=pres.addSlide();slide.background={color:NAVY};
  addSlideHeader(slide,'07 · Next Best Actions',pres);
  slide.addText('Next Best Actions',{x:0.3,y:0.52,w:9.4,h:0.5,fontSize:22,fontFace:'Calibri',bold:true,color:'FFFFFF',margin:0});
  const coaching=opp.coaching,actions=(coaching&&coaching.actions)||[],stage=(coaching&&coaching.stage)||'mid';
  const stageMap={early:'Early stage',mid:'Mid stage',late:'Late stage'};
  slide.addText(stageMap[stage]||'Mid stage',{x:0.3,y:1.02,w:9.4,h:0.28,fontSize:10,fontFace:'Calibri',color:TEAL,margin:0});
  if(!actions.length){slide.addText('No coaching actions generated yet.',{x:0.3,y:1.8,w:9.4,h:0.4,fontSize:11,fontFace:'Calibri',color:SLATE,italic:true,margin:0});return;}
  actions.slice(0,3).forEach((a,i)=>{
    const y=1.45+i*1.3;
    slide.addShape(pres.shapes.RECTANGLE,{x:0.3,y,w:9.4,h:1.12,fill:{color:'162030'},line:{color:'1f3040',width:0.5},shadow:makeShadow()});
    slide.addShape(pres.shapes.RECTANGLE,{x:0.3,y,w:0.06,h:1.12,fill:{color:TEAL}});
    slide.addText(String(i+1),{x:0.45,y:y+0.08,w:0.4,h:0.45,fontSize:26,fontFace:'Calibri',bold:true,color:TEAL,margin:0});
    slide.addText(trunc(a.action,100),{x:0.95,y:y+0.1,w:8.5,h:0.42,fontSize:13,fontFace:'Calibri',bold:true,color:'FFFFFF',wrap:true,margin:0});
    slide.addText(trunc(a.why,180),{x:0.95,y:y+0.56,w:7.8,h:0.46,fontSize:9.5,fontFace:'Calibri',color:SLATE,wrap:true,italic:true,margin:0});
    if(a.section)slide.addText(a.section.replace('_',' ').toUpperCase(),{x:8.6,y:y+0.78,w:1.0,h:0.22,fontSize:7,fontFace:'Calibri',bold:true,color:TEAL,align:'right',charSpacing:1,margin:0});
  });
  slide.addText('Generated by Playbk Labs · playbklabs.ai',{x:0.3,y:5.35,w:9.4,h:0.2,fontSize:7.5,fontFace:'Calibri',color:DIM,align:'center',margin:0});
}

// ── Main handler ──────────────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const opp = req.body;
    if (!opp || !opp.name) {
      return res.status(400).json({ error: 'Invalid opportunity data' });
    }

    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';
    pres.author = 'Playbk Labs';
    pres.title  = `Deal Review — ${clean(opp.name)}`;

    const MEDD = [
      { id:'economic_buyer',   title:'Economic Buyer',      blocker:true,  questions:Array.from({length:4}) },
      { id:'champion',         title:'Champion',            blocker:true,  questions:Array.from({length:4}) },
      { id:'paper_process',    title:'Paper Process',       blocker:true,  questions:Array.from({length:4}) },
      { id:'decision_process', title:'Decision Process',    blocker:true,  questions:Array.from({length:4}) },
      { id:'data_integration', title:'Data & Integration',  blocker:true,  questions:Array.from({length:4}) },
      { id:'metrics',          title:'Metrics',             blocker:false, questions:Array.from({length:4}) },
      { id:'decision_criteria',title:'Decision Criteria',   blocker:false, questions:Array.from({length:4}) },
      { id:'competition',      title:'Competition',         blocker:false, questions:Array.from({length:4}) },
    ];

    slideCover         (pres, opp);
    slideSituation     (pres, opp);
    slideStakeholders  (pres, opp);
    slideDeal          (pres, opp);
    slideQualification (pres, opp, MEDD);
    slideMAP           (pres, opp);
    slideBizAsks       (pres, opp);
    slideCoaching      (pres, opp);

    const buffer = await pres.write({ outputType: 'nodebuffer' });
    const filename = (opp.name || 'deal-review').replace(/[^a-z0-9]/gi, '_') + '_deal_review.pptx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.status(200).send(buffer);

  } catch (e) {
    console.error('Deck generation error:', e);
    res.status(500).json({ error: e.message });
  }
}
