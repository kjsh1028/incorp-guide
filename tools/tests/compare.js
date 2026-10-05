// 화면(index.html)의 문서 만들기·회신 읽기가 tools/docgen.py 와 같은 결과를 내는지 비교한다.
// 가상 값만 쓴다(tools/sample_case.json, 가상 회신). 고객 자료를 넣지 않는다.
//   NODE_PATH=<playwright 위치> node tools/tests/compare.js
// 필요: python3 + python-docx, Node + playwright(Chromium)
const { chromium } = require('playwright'); const fs=require('fs'), os=require('os'), path=require('path'), cp=require('child_process');
const R=path.resolve(__dirname,'..','..')+'/', W=fs.mkdtempSync(path.join(os.tmpdir(),'cmp-'));
const py=(...a)=>cp.execFileSync('python3',a,{cwd:R,encoding:'utf8'});
let bad=0; const ok=(name,c)=>{console.log((c?'같음  ':'다름! ')+name); if(!c)bad++};
(async () => {
  const html=fs.readFileSync(R+'index.html','utf8'), cut=(a,b)=>html.slice(html.indexOf(a),html.indexOf(b));
  const src=cut('/* ---------- 고객 문서 만들기 ----------','function loadScript(')+cut('/* ---------- 고객 회신 읽기 ----------','/* 사건 기록에는')
    +'\nwindow.T_BUILD=buildXml;window.T_SPLAN=signPlan;window.T_SFILL=signFill;window.T_READ=function(x,so){return readReply(readTables(x),so)};';
  const br=await chromium.launch(), p=await br.newPage(); await p.setContent('<html><body></body></html>');
  await p.addScriptTag({path:R+'lib/jszip.min.js'}); await p.addScriptTag({content:src});
  const xmlOf=f=>p.evaluate(async b=>{const z=await JSZip.loadAsync(Uint8Array.from(atob(b),c=>c.charCodeAt(0)));return z.file('word/document.xml').async('string')},fs.readFileSync(f).toString('base64'));
  const firm=JSON.parse(fs.readFileSync(R+'tools/firm_profile.json','utf8'));
  // 1) 문서 3종: 샘플 사건과, 이사·감사가 있고 보고가 빈 사건
  const c1=JSON.parse(fs.readFileSync(R+'tools/sample_case.json','utf8'));
  const c2=JSON.parse(JSON.stringify(c1)); c2.officers={other_directors:1,auditor:true}; c2.investor_options=[];
  c2.report={no:3,done:[],requests:[],next_plan:{ko:'시험',zh:'测试'},schedule:[{when:{ko:'완료',zh:'已完成'},status:'done'},{when:{ko:'해당 없음',zh:'不适用'}}]};
  const c3=JSON.parse(JSON.stringify(c1)); c3.investor_options=[{label:{ko:'홍콩 법인',zh:'香港公司'},type:'hold'},{label:{ko:'싱가포르 법인',zh:'新加坡公司'},type:'hold'}];
  for (const [tag,c] of [['샘플',c1],['변형',c2],['후보둘',c3]]) {
    const cf=W+'/'+tag+'.json', out=W+'/out_'+tag; fs.writeFileSync(cf,JSON.stringify(c));
    const d={client_cn:c.client_name_cn,date:c.date,reply_by:c.reply_by,alts:c.investor_options.map(o=>o.label),officers:c.officers,report:c.report};
    for (const kind of ['request','signing','progress']) {
      const made=py('tools/docgen.py',kind,cf,out).trim().split('\n').map(l=>l.split(' ')[0]);
      for (const f of made) {
        const lang=/_KR_/.test(f)?'ko':'zh', tpl=await xmlOf(R+'tools/templates/'+kind+'_'+lang+'.docx');
        const js=await p.evaluate(([x,k,d,firm,l])=>window.T_BUILD(k,x,d,firm,l).xml,[tpl,kind,d,firm,lang]);
        ok(`문서 ${tag} ${kind} ${lang}`, js===await xmlOf(f));
      }
    }
  }
  // 1-2) 서명 서류 A1~A6: 샘플(대표이사만, 율촌 값 비움)과 이사 2명·감사가 있고 율촌 값이 있는 사건
  const od=JSON.parse(fs.readFileSync(R+'data/tracks/KR-JSC/outputs.json','utf8')).documents.find(d=>d.id==='sign_forms');
  const s2=JSON.parse(JSON.stringify(c1)); const pp=n=>({name_en:'TEST '+n,nationality_en:'People\'s Republic of China',dob_en:'May 5, 1985',addr_en:n+' Test Road'});
  s2.sign.people.director=[pp('D1'),pp('D2')]; s2.sign.people.auditor=[pp('AU')]; s2.sign.values.head_office_addr_en='';
  s2.firm={...firm,sign:{yulchon_attorneys:'Test Attorney A and Test Attorney B',seal_attorney_en:'Test Attorney A',seal_attorney_reg_no:'00000'}};
  for (const [tag,c] of [['샘플',c1],['이사·감사',s2]]) {
    const cf=W+'/s_'+tag+'.json', out=W+'/sout_'+tag; fs.writeFileSync(cf,JSON.stringify(c));
    const made=py('tools/docgen.py','sign',cf,out).trim().split('\n').map(l=>l.split(' | ')[0].trim());
    const plan=await p.evaluate(([od,sd,f,sh])=>window.T_SPLAN(od,sd,f,sh),[od,c.sign,c.firm||firm,c.client_short]);
    ok(`서명 서류 ${tag} 개수 ${made.length}`, plan.length===made.length);
    for (let i=0;i<made.length;i++) {
      const tpl=await xmlOf(R+plan[i].form.template);
      const js=await p.evaluate(([x,v])=>window.T_SFILL(x,v).xml,[tpl,plan[i].vals]);
      ok(`서명 서류 ${tag} ${path.basename(made[i])}`, js===await xmlOf(made[i]) && path.basename(made[i])===plan[i].name.replace(/[\\/:*?"<>|]/g,'_'));
    }
  }
  // 2) 회신 읽기: 가상 회신(국문 채움)과 빈 회신
  py('tools/docgen.py','request','tools/sample_case.json',W+'/req');
  const ko=fs.readdirSync(W+'/req').find(n=>/_KR_/.test(n));
  py('tools/tests/make_fake_reply.py',W+'/req',W+'/fake_reply.docx');
  const so=[{label:c1.client_name_cn,type:'cn'},{label:c1.investor_options[0].label,type:c1.investor_options[0].type}];
  for (const [tag,f] of [['가상 회신',W+'/fake_reply.docx'],['빈 회신',W+'/req/'+ko]]) {
    const pyr=JSON.parse(py('tools/docgen.py','read','tools/sample_case.json',f));
    const js=await p.evaluate(([x,so])=>window.T_READ(x,so),[await xmlOf(f),so]);
    ok('회신 '+tag, ['answers','program_inputs','review'].every(k=>JSON.stringify(js[k])===JSON.stringify(pyr[k])));
  }
  await br.close(); fs.rmSync(W,{recursive:true,force:true});
  console.log(bad?`다른 것 ${bad}개`:'모두 같음'); process.exit(bad?1:0);
})();
