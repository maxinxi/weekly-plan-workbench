/* Optional end-to-end check: npm install --no-save playwright, then node tests/offline_ui.cjs.
   Uses an installed Edge (or PLAYWRIGHT_BROWSER executable). No server or network is used. */
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const {pathToFileURL}=require('node:url');
const {chromium}=require('playwright');
async function main() {
  const browser=await chromium.launch({headless:true,...(process.env.PLAYWRIGHT_BROWSER?
    {executablePath:process.env.PLAYWRIGHT_BROWSER}:{channel:'msedge'})});
  try {
    const context=await browser.newContext({acceptDownloads:true,offline:true});
    const page=await context.newPage();
    const errors=[],requests=[];
    page.on('pageerror',e=>errors.push(e.message));
    page.on('console',m=>{if(m.type()==='error') console.error(m.text());});
    page.on('request',r=>{if(/^https?:/.test(r.url()))requests.push(r.url());});
    await page.goto(pathToFileURL(path.resolve(__dirname,'../周计划工作台.html')).href);
    let bytes,name;
    if(process.argv[2]) {bytes=fs.readFileSync(process.argv[2]);name=path.basename(process.argv[2]);}
    else {bytes=Buffer.from(await page.evaluate(async()=>Array.from(new Uint8Array(await WeeklyPlan.writeSampleSourceXlsx()))));name='离线测试（第36周）.xlsx';}
    await page.locator('#file').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});
    await page.waitForFunction(()=>!document.getElementById('btnProcessedSrc').disabled,{},{timeout:60000});
    const state=await page.locator('#sourceStatus').innerText();
    assert(state.includes('源表已整理'),state);
    assert.equal(await page.evaluate(()=>WeeklyPlan.sourceProcessor.work.size),0);
    const report=await page.evaluate(()=>sourceReport);
    assert(!report.issues.some(i=>i.code==='BACKUP_MEMORY_ONLY'),'IndexedDB backup should work offline');
    assert.deepEqual(report.stages,['backup_work','preprocess','width','height','insert','output','work_cleanup']);
    assert.equal(await page.locator('#btnOutputDirectory').count(),0,'no directory selection');
    const download=page.waitForEvent('download');await page.locator('#btnExportAll').click();
    const file=await download.catch(async e=>{console.error(await page.evaluate(()=>({issues:sourceReport.issues.filter(i=>i.level==='error'),result:!!result,exporting})));throw e;});
    assert.equal(file.suggestedFilename(),'（处理后）.zip');
    assert.equal(await file.failure(),null);
    if(process.argv[3]) {
      fs.mkdirSync(process.argv[3],{recursive:true});
      await file.saveAs(path.join(process.argv[3],file.suggestedFilename()));
      fs.writeFileSync(path.join(process.argv[3],'source.报告.json'),JSON.stringify(report,null,2));
    }
    // Reload proves the original backup lives in IndexedDB, not a JavaScript variable.
    await page.reload();
    const saved=await page.evaluate(async filename=> {
      const sp=WeeklyPlan.sourceProcessor;
      await sp.backupOnce(filename,new Uint8Array([1,2,3]).buffer,sp.report(filename));
      return Array.from(new Uint8Array(await sp.getBackup(filename)));
    },name);
    assert.equal(Buffer.compare(Buffer.from(saved),bytes),0,'existing backup must not be overwritten');
    assert.deepEqual(requests,[],'single HTML must not request external assets');
    assert.deepEqual(errors,[],'browser console errors');
    console.log('Offline Edge UI passed: import, download, persistent first backup, cleanup, zero network requests.');
    console.log(state);
  } finally {await browser.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});
