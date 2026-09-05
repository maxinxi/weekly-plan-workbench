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
    assert.equal(downloaded.suggestedFilename(),'启动日计划工作台.cmd');await downloaded.saveAs(path.join(out,'launcher.cmd'));
    const cmd=fs.readFileSync(path.join(out,'launcher.cmd'),'utf8');
    const data=JSON.parse(Buffer.from(cmd.split('__DAILY_DATA__\r\n')[1].trim(),'base64').toString('utf8'));
    assert(data['daily_plan_legacy.py'].includes('def preprocess_excel'));
    assert(data['日计划工作台.html'].includes('日计划工作台'));
    const ps=Buffer.from(cmd.match(/-EncodedCommand ([A-Za-z0-9+/=]+)/)[1],'base64').toString('utf16le');
    fs.writeFileSync(path.join(out,'launcher.ps1'),ps,'utf8');
    assert(ps.includes('-WindowStyle Hidden'));assert(ps.includes('LastIndexOf'));
    assert.deepEqual(errors,[]);await offline.close();
    proc=spawn(python,['-X','utf8','daily_plan_server.py','--no-browser','--base',path.join(out,'server')],{cwd:root,windowsHide:true,stdio:['ignore','pipe','pipe']});
    const url=await new Promise((resolve,reject)=>{let output='';const timer=setTimeout(()=>reject(Error('Server startup timeout')),30000);proc.stdout.on('data',x=>{output+=x.toString('utf8');const m=output.match(/http:\/\/127\.0\.0\.1:\d+\/\?token=[\w-]+/);if(m){clearTimeout(timer);resolve(m[0]);}});proc.on('exit',code=>reject(Error('Server exited '+code)));});
    const ctx=await browser.newContext({acceptDownloads:true});const p=await ctx.newPage();const pageErrors=[];p.on('pageerror',e=>pageErrors.push(e.message));
    let external=0;await ctx.route('**/*',route=>{if(/^https?:/.test(route.request().url())&&!route.request().url().startsWith('http://127.0.0.1:')){external++;return route.abort();}return route.continue();});
    await p.goto(url);await p.locator('#importPanel').waitFor({state:'visible'});
    await p.locator('#importPanel summary').click();await p.locator('#files').setInputFiles([path.join(out,'9月7日现场.xlsx'),path.join(out,'9月7日风险管控.xlsx')]);await p.locator('#import').click();
    await p.locator('.job').first().waitFor({timeout:30000});assert.equal(await p.locator('.job').count(),2);
    const job=p.locator('.job').first();await job.locator('[data-choice=am]').selectOption('0');await job.locator('[data-choice=pm]').selectOption('1');await job.locator('[data-write]').click();
    await p.waitForFunction(()=>document.querySelector('[data-same]').value.includes('下午：'));
    const val=await p.locator('[data-same]').first().inputValue();assert(val.includes('上午：测试甲'));assert(val.includes('下午：测试乙'));
    await p.locator('.job').nth(1).locator('[data-same]').fill('测试丙');await p.locator('.job').nth(1).locator('[data-same]').blur();
    await p.locator('#confirmDialog').waitFor({state:'visible'});await p.locator('[data-close=confirmDialog]').click();
    await p.locator('#summaryA').click();await p.waitForFunction(()=>document.getElementById('textA').value.includes('【完整版】'));
    assert.equal(await p.locator('#textB').inputValue(),'');
    await p.reload();await p.waitForFunction(()=>document.querySelectorAll('.job').length===2);
    assert((await p.locator('[data-same]').first().inputValue()).includes('下午：测试乙'));
    await p.locator('#summaryA').click();await p.waitForFunction(()=>document.getElementById('textA').value.includes('【完整版】'));
    await p.screenshot({path:path.join(out,'workbench.png'),fullPage:true});
    // An untrusted origin cannot trigger local processing even with a guessed port.
    const blocked=await p.evaluate(async()=>{const r=await fetch('/api/update',{method:'POST',headers:{'Content-Type':'application/json'},body:'{"edits":[]}'});return r.status;});assert.equal(blocked,403);
    assert.equal(external,0);assert.deepEqual(pageErrors,[]);
    console.log('Offline launcher + local import + per-row dropdowns + morning/afternoon + Mode A + readiness dialog + local authorization passed');
  }finally{if(proc)proc.kill();await browser.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1});
