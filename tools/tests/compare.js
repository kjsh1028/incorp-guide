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
    +'\nwindow.T_BUILD=buildXml;window.T_READ=function(x,so){return readReply(readTables(x),so)};';
  const br=await chromium.launch(), p=await br.newPage(); await p.setContent('<html><body></body></html>');
  await p.addScriptTag({path:R+'lib/jszip.min.js'}); await p.addScriptTag({content:src});
  const xmlOf=f=>p.evaluate(async b=>{const z=await JSZip.loadAsync(Uint8Array.from(atob(b),c=>c.charCodeAt(0)));return z.file('word/document.xml').async('string')},fs.readFileSync(f).toString('base64'));
  const firm=JSON.parse(fs.readFileSync(R+'tools/firm_profile.json','utf8'));
  // 1) 문서 3종: 샘플 사건과, 이사·감사가 있고 보고가 빈 사건
  const c1=JSON.parse(fs.readFileSync(R+'tools/sample_case.json','utf8'));
  const c2=JSON.parse(JSON.stringify(c1)); c2.officers={other_directors:1,auditor:true}; c2.investor_options=[];
  c2.report={no:3,done:[],requests:[],next_plan:{ko:'시험',zh:'测试'},schedule:[{when:{ko:'완료',zh:'已完成'},status:'done'},{when:{ko:'해당 없음',zh:'不适用'}}]};
  for (const [tag,c] of [['샘플',c1],['변형',c2]]) {
    const cf=W+'/'+tag+'.json', out=W+'/out_'+tag; fs.writeFileSync(cf,JSON.stringify(c));
    const d={client_cn:c.client_name_cn,date:c.date,reply_by:c.reply_by,alt:(c.investor_options[0]||{}).label||{},officers:c.officers,report:c.report};
    for (const kind of ['request','signing','progress']) {
      const made=py('tools/docgen.py',kind,cf,out).trim().split('\n').map(l=>l.split(' ')[0]);
      for (const f of made) {
        const lang=/_KR_/.test(f)?'ko':'zh', tpl=await xmlOf(R+'tools/templates/'+kind+'_'+lang+'.docx');
        const js=await p.evaluate(([x,k,d,firm,l])=>window.T_BUILD(k,x,d,firm,l).xml,[tpl,kind,d,firm,lang]);
        ok(`문서 ${tag} ${kind} ${lang}`, js===await xmlOf(f));
      }
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
