/* Test the exact ExcelJS bundled in the shipped HTML. No npm download is needed. */
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const root=path.resolve(__dirname,'..');
const html=fs.readFileSync(path.join(root,'周计划工作台.html'),'utf8');
const bundle=[...html.matchAll(/<script>([\s\S]*?)<\/script>/g)][0][1];
const ctx={console,setTimeout,clearTimeout,setInterval,clearInterval,TextEncoder,TextDecoder,Buffer,Uint8Array,ArrayBuffer,Date,Blob,URL};
ctx.window=ctx;ctx.self=ctx;ctx.global=ctx;
vm.createContext(ctx);vm.runInContext(bundle+'\nthis.WP=WeeklyPlan;',ctx);
const WP=ctx.WP, SP=WP.sourceProcessor;
// The exposed factory accepts the already loaded library; use the bundled sample writer for load coverage.
async function main() {
  assert.equal(SP.outputName('表1（第36周）.xlsx'),'（第36周）（处理后的源表）.xlsx');
  assert.equal(SP.outputName('源表.xlsx'),'源表（处理后的源表）.xlsx');
  for(const name of ['~$表1.xlsx','明细.xlsx','处理后的源表.xlsx','work/表1.xlsx']) assert(!SP.isSource(name));
  for(const [value,code] of [['2026/2/30 08:30-17:30','DATE_INVALID'],['08:30-17:30','TIME_DATE_MISSING'],['2026/9/1 25:00-26:00','TIME_INVALID'],['2026/9/1 17:30-08:30','TIME_ORDER'],['2026/9/1 上午','TIME_FORMAT']]) {
    const log=SP.report('测试');const [out]=SP.validateTime(value,log,'D4');
    assert.equal(out,value);assert(log.issues.some(i=>i.code===code));
  }
  const sample=await WP.writeSampleSourceXlsx();
  const parsed=await WP.parseSourceWorkbook(sample,'样例.xlsx');
  const index=parsed.headers.findIndex(h=>String(h).includes('管控措施'));
  assert(index>=0);
  for(const row of parsed.rows) row[index]='ONLY_SOURCE_MEASURES_SENTINEL';
  const detail=WP.runPipeline({...parsed,sourceName:'样例.xlsx',measuresPlus:1});
  assert(!JSON.stringify(detail.layout).includes('ONLY_SOURCE_MEASURES_SENTINEL'),'weekly detail must exclude measures');
  const output=await SP.processBuffer(sample,'样例（第36周）.xlsx',1);
  assert(output.buffer.byteLength>1000);
  assert.deepEqual(Array.from(output.report.stages),['preprocess','width','height','insert','output']);
  if(process.argv[2]==='--sample') {fs.writeFileSync(process.argv[3],Buffer.from(output.buffer));return;}
  const input=process.argv[2];
  if(input) {
    const bytes=fs.readFileSync(input), buffer=bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength);
    const before=Buffer.from(buffer);
    const filename=path.basename(input);
    const out=await SP.processFile({name:filename,arrayBuffer:async()=>buffer},1);
    assert.equal(Buffer.compare(before,Buffer.from(buffer)),0);
    assert.equal(SP.work.size,0);
    const backup=Buffer.from(await SP.getBackup(filename));
    await SP.backupOnce(filename,new Uint8Array([1,2,3]).buffer,SP.report(filename));
    assert.equal(Buffer.compare(backup,Buffer.from(await SP.getBackup(filename))),0);
    const directory=process.argv[3]||path.dirname(input);
    fs.mkdirSync(directory,{recursive:true});
    fs.writeFileSync(path.join(directory,'browser.xlsx'),Buffer.from(out.buffer));
    fs.writeFileSync(path.join(directory,'browser.report.json'),JSON.stringify(out.report,null,2));
    fs.writeFileSync(path.join(directory,'backup.zip'),Buffer.from(await SP.backupZip(filename)));
    console.log(JSON.stringify({inserted:out.inserted,deleted:out.deletedRows,issues:out.report.issues.map(i=>i.code),stages:out.report.stages}));
  }
  try {await SP.processFile({name:'invalid.xlsx',arrayBuffer:async()=>new Uint8Array([1,2,3]).buffer});assert.fail('invalid input accepted');} catch(e) {assert.equal(SP.work.size,0);}
  console.log('Browser core tests passed (exact embedded bundle)');
}
main().catch(e=>{console.error(e);process.exitCode=1;});
