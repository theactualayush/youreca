/* Checks engine.js against values cached in FF_Meeting_Calculator_OptimizeTest.xlsx. Run: node tests/verify-against-excel.js */
const fs=require('fs'),path=require('path'); global.window=globalThis;
eval(fs.readFileSync(path.join(__dirname,'..','engine.js'),'utf8'));
eval(fs.readFileSync(path.join(__dirname,'..','seed.js'),'utf8'));
const D=JSON.parse(fs.readFileSync(path.join(__dirname,'excel-expected.json'),'utf8'));
const E=FFEngine,U=E.util,S=window.FF_SEED;
const inp={today:U.fromText(S.asof),confirmed:S.confirmed.map(U.fromText),mkt:S.mkt,
 sofr:S.sofr.map(r=>[U.fromText(r[0]),r[1]]),effr:S.effr.map(r=>[U.fromText(r[0]),r[1]]),holidays:S.holidays.map(U.fromText),
 cases:{rows:S.cases.rows,extra:S.cases.extra.map(r=>({...r,date:U.fromText(r.date)})),cutoff:U.fromText(S.cases.cutoff)}};
const t0=Date.now(); const R=E.run(inp,S.scenarios[S.selected].map(U.fromText)); const ms=Date.now()-t0;
let ok=0,bad=0; const cmp=(name,a,b)=>{const be=(typeof b==='string'||b===null), ae=(typeof a!=='number'||isNaN(a)); if(be||ae){ if(be&&ae)ok++; else {bad++;console.log('ERRMISMATCH',name,a,b)} return;} if(Math.abs(a-b)<1e-8)ok++; else {bad++;console.log('DIFF',name,a,b,a-b)}};
R.rows.forEach((r,i)=>cmp('cut '+r.month,r.cut,D.meetingCuts[i]['Cuts/Hikes']));
D.SR3SR1_Fair.forEach((row,k)=>{const f=R.fair.get(row.Contract); if(!f){console.log('missing',row.Contract);bad++;return}
 cmp('sr3 '+row.Contract,f.sr3,row['SR3 Fair']);cmp('sr1 '+row.Contract,f.sr1,row['SR1 Fair']);cmp('zq '+row.Contract,f.zq,row['ZQ Fair']);
 for(const j of [0,1]){cmp('sr3c'+j+row.Contract,f.sr3c[j],D.SR3_Cases[k]['Case '+(j+1)]);cmp('sr1c'+j+row.Contract,f.sr1c[j],D.SR1_Cases[k]['Case '+(j+1)]);cmp('zqc'+j+row.Contract,f.zqc[j],D.ZQ_Cases[k]['Case '+(j+1)]);}});
const s=R.src,g=D.graphs_cached;
g.ZQ_structures.forEach(r=>{const l=r['ZQ Structures'];cmp('zqs H1 '+l,E.binom(l,s.zqCase(0)),r['Hike 1']);cmp('zqs H2 '+l,E.binom(l,s.zqCase(1)),r['Hike 2']);cmp('zqs live '+l,E.binom(l,s.zqLive),r.Live);cmp('zqs fair '+l,E.binom(l,s.zqFair),r.Fair)});
g.SR3SR1_structures.forEach(r=>{const l=r['SR3-SR1 Structure'];cmp('s31 fair '+l,E.flyX(l,s.sr3Fair,s.sr1Fair),r.Fair);cmp('s31 c1 '+l,E.flyX(l,s.sr3Case(0),s.sr1Case(0)),r['Case 1']);cmp('s31 c2 '+l,E.flyX(l,s.sr3Case(1),s.sr1Case(1)),r['Case 2']);if(r.VWAP)cmp('s31 vwap '+l,E.flyX(l,s.sr3Live,s.sr1Live),r.VWAP);if(r.Settle)cmp('s31 set '+l,E.flyX(l,s.sr3Settle,s.sr1Settle),r.Settle)});
g.SR3_ZQ.forEach(r=>{const l=r['SR3-ZQ'];cmp('s3z fair '+l,E.flyX(l,s.sr3Fair,s.zqFair),r.Fair);cmp('s3z h1 '+l,E.flyX(l,s.sr3Case(0),s.zqCase(0)),r['Hike 1']);cmp('s3z live '+l,E.flyX(l,s.sr3Live,s.zqLive),r.Live)});
g.SR3_ZQ_4leg.forEach(r=>{const l=r['SR3-ZQ 4 Leg'];cmp('4l fair '+l,E.four(l,s.sr3Fair,s.zqFair),r.Fair);cmp('4l h2 '+l,E.four(l,s.sr3Case(1),s.zqCase(1)),r['Hike 2']);cmp('4l live '+l,E.four(l,s.sr3Live,s.zqLive),r.Live)});
g.SR3_flys.forEach(r=>{const l=r['SR3 Flys'];cmp('sf price '+l,E.binom(l,s.sr3Live),r.Price);cmp('sf fair '+l,E.binom(l,s.sr3Fair),r.Fair);cmp('sf h1 '+l,E.binom(l,s.sr3Case(0)),r['Hike 1']);cmp('sf set '+l,E.binom(l,s.sr3Settle),r.Settle)});
g.ff_1mo_flys.forEach(r=>{const l=r['Fed Fund 1mo Flys'];cmp('ff price '+l,E.binom(l,s.zqLive),r.Price);cmp('ff fair '+l,E.binom(l,s.zqFair),r.Fair);cmp('ff h2 '+l,E.binom(l,s.zqCase(1)),r['Hike 2'])});
// meeting neutrals
g.meeting_neutrals.forEach((r,i)=>{const m=R.mn[i]; if(!m){console.log('no mn',i);bad++;return} const lbl=m.legs.join('-'); if(lbl!==r['Meeting Neutrals'])console.log('MN label',lbl,r['Meeting Neutrals']);
 cmp('mn settle '+lbl,100*E.ratioSum(m.legs,m.ratios,s.zqSettle),r.Settle);
 const sp=[m.ratios[0],m.ratios[0]+m.ratios[1],0]; const v=m.legs.reduce((a,l,j)=>a+sp[j]*(j<2?s.zq1ms(l):0),0); cmp('mn vwappnl '+lbl,v*20.48*2,r['VWAP PnL'])});
cmp('P8',R.P8,D.Graphs_P8_total_hikes_2027);
console.log('ok',ok,'bad',bad,'ms',ms);
process.exitCode = bad ? 1 : 0;
