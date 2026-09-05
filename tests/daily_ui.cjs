/* End-to-end local UI / launcher tests. Python & Excel dependencies must be installed.
   PYTHON_EXE selects that installation. No remote website is contacted. */
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const {spawn,execFileSync}=require('node:child_process');
const {pathToFileURL}=require('node:url');
const {chromium}=require('playwright');
const root=path.resolve(__dirname,'..'),python=process.env.PYTHON_EXE||'python';
const out=path.join(root,'work','daily-ui');fs.mkdirSync(out,{recursive:true});
async function main(){
  execFileSync(python,['-X','utf8','-c',"import sys;from pathlib import Path;sys.path.insert(0,'tests');from test_daily_plan import sample;[(Path('work/daily-ui')/('9月7日'+('现场' if k=='site' else '风险管控')+'.xlsx')).write_bytes(sample(k)) for k in ('site','risk')]"],{cwd:root,windowsHide:true});
  const browser=await chromium.launch({headless:true,channel:'msedge'});
  let proc;
  try{
    const offline=await browser.newContext({offline:true,acceptDownloads:true});
    const page=await offline.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.goto(pathToFileURL(path.join(root,'日计划工作台.html')).href);
    const pending=page.waitForEvent('download');await page.locator('#launch').click();const downloaded=await pending;
    assert.equal(downloaded.suggestedFilename(),'启动日计划工作台.py');await downloaded.saveAs(path.join(out,'launcher.py'));
    const launcher=fs.readFileSync(path.join(out,'launcher.py'),'utf8');
    assert(launcher.startsWith('#! python3'));
    assert(launcher.includes('日计划工作台：双击启动'));
    assert(launcher.includes("subprocess.run([sys.executable"));
    assert(!launcher.includes('EncodedCommand'));assert(!launcher.includes('__DAILY_DATA__'));
    assert.deepEqual(errors,[]);await offline.close();
    proc=spawn(python,['-X','utf8','daily_plan_server.py','--no-browser','--base',path.join(out,'server')],{cwd:root,windowsHide:true,stdio:['ignore','pipe','pipe']});
    const url=await new Promise((resolve,reject)=>{let output='';const timer=setTimeout(()=>reject(Error('Server startup timeout')),30000);proc.stdout.on('data',x=>{output+=x.toString('utf8');const m=output.match(/http:\/\/127\.0\.0\.1:\d+\/\?token=[\w-]+/);if(m){clearTimeout(timer);resolve(m[0]);}});proc.on('exit',code=>reject(Error('Server exited '+code)));});
    const ctx=await browser.newContext({acceptDownloads:true});const p=await ctx.newPage();const pageErrors=[];p.on('pageerror',e=>pageErrors.push(e.message));
    let external=0;await ctx.route('**/*',route=>{if(/^https?:/.test(route.request().url())&&!route.request().url().startsWith('http://127.0.0.1:')){external++;return route.abort();}return route.continue();});
    await p.goto(url);await p.locator('#importPanel').waitFor({state:'visible'});
    const dropped=[['9月7日现场.xlsx',fs.readFileSync(path.join(out,'9月7日现场.xlsx')).toString('base64')],['9月7日风险管控.xlsx',fs.readFileSync(path.join(out,'9月7日风险管控.xlsx')).toString('base64')]];
    await p.evaluate(items=>{const dt=new DataTransfer();for(const [name,b64] of items){const raw=atob(b64),bytes=Uint8Array.from(raw,c=>c.charCodeAt(0));dt.items.add(new File([bytes],name,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));}document.getElementById('dropzone').dispatchEvent(new DragEvent('drop',{bubbles:true,dataTransfer:dt}));},dropped);
    assert.equal(await p.locator('#import').isEnabled(),true);await p.locator('#import').click();
    await p.locator('.job').first().waitFor({timeout:30000});assert.equal(await p.locator('.job').count(),2);
    assert.match(await p.locator('#nextStep').textContent(),/正式文件尚未生成.*剩余 1 项/s);
    assert.equal(await p.locator('[data-same]').count(),1);assert.match(await p.locator('.job').nth(1).textContent(),/管理人员.*不需要领导选择/s);
    const job=p.locator('.job').first();await job.locator('[data-choice=am]').selectOption('0');await job.locator('[data-choice=pm]').selectOption('1');await job.locator('[data-write]').click();
    await p.locator('#confirmDialog').waitFor({state:'visible'});await p.locator('[data-close=confirmDialog]').click();
    assert.match(await p.locator('.job').first().textContent(),/上午：测试甲[\s\S]*下午：测试乙/);assert.match(await p.locator('#progress').textContent(),/1 \/ 1/);
    await p.locator('#summaryA').click();await p.waitForFunction(()=>document.getElementById('textA').value.includes('【完整版】'));
    assert.equal(await p.locator('#textB').inputValue(),'');
    await p.reload();await p.waitForFunction(()=>document.querySelectorAll('.job').length===2);
    assert.equal(await p.locator('[data-clear]').count(),1);assert.match(await p.locator('.job').first().textContent(),/下午：测试乙/);
    assert.match(await p.locator('#nextStep').textContent(),/人员选择已完成.*生成到岗到位/s);
    await p.locator('#summaryA').click();await p.waitForFunction(()=>document.getElementById('textA').value.includes('【完整版】'));
    await p.evaluate(async()=>setModel(await api('update',{edits:[{kind:'risk',row:4,field:'work',value:'措辞不同的设备作业'}]})));
    assert.match(await p.locator('#alerts').textContent(),/待人工判断.*整理后第 1 项（原表第 4 行）/s);
    await p.locator('[data-issue]').first().click();assert.equal(await p.locator('#fixCompare .compare-value').count(),2);assert.equal(await p.locator('#fixChoices button').count(),2);await p.locator('[data-close=fixDialog]').click();
    const fakeResult={folder:'C:/测试源表/（处理后）',files:['现场（到岗到位）.pdf','计划（风险管控）.pdf'],summaries:{a:'模式 A',b:'模式 B'},warnings:[]};
    await p.route('**/api/status',route=>route.fulfill({json:{status:{busy:false,phase:'完成',log:[],error:''},result:fakeResult}}));
    await p.evaluate(()=>poll());
    assert.equal(await p.locator('dialog[open]').count(),0,'generation must not open print preview');
    assert.match(await p.locator('#outputLocation').textContent(),/测试源表.*（处理后）/);
    await p.route('**/api/file?**',route=>route.fulfill({contentType:'text/plain',body:'PDF fixture'}));
    await p.locator('#preview').click();
    assert.equal(await p.locator('#pdfDialog').isVisible(),true);
    assert.equal(await p.locator('#pdfSelect option').count(),2);
    await p.locator('#pdfSelect').selectOption('计划（风险管控）.pdf');
    assert.match(decodeURIComponent(await p.locator('#pdfDownload').getAttribute('href')),/风险管控/);
    await p.locator('[data-close=pdfDialog]').click();
    await p.screenshot({path:path.join(out,'workbench.png'),fullPage:true});
    // An untrusted origin cannot trigger local processing even with a guessed port.
    const blocked=await p.evaluate(async()=>{const r=await fetch('/api/update',{method:'POST',headers:{'Content-Type':'application/json'},body:'{"edits":[]}'});return r.status;});assert.equal(blocked,403);
    assert.equal(external,0);assert.deepEqual(pageErrors,[]);
    console.log('Offline launcher + drag import + leader-only dropdown + order-independent cleanup audit + morning/afternoon + Mode A + readiness dialog + local authorization passed');
  }finally{if(proc)proc.kill();await browser.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1});
