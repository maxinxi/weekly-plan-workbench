/* Browser source processing. Embedded verbatim by tools/build_html.py; no network dependencies. */
function createSourceProcessor(ExcelJS) {
  "use strict";
  const MAX_HEIGHT = 409.5;
  const KEYS = {
    content: ["作业内容", "工作内容", "作业项目", "作业任务"],
    time: ["作业时间", "工作时间", "计划时间", "起止时间", "时间"],
    measures: ["管控措施", "控制措施", "防控措施", "风险措施"],
    person: ["同进同出", "负责人", "人员", "联系人", "姓名"], date: ["日期"], risk:["作业风险等级","风险等级","风险级别"]
  };
  const DATE = "(\\d{4})[年/.-](\\d{1,2})[月/.-](\\d{1,2})日?";
  const CLOCK = "(\\d{1,2}):(\\d{2})";
  const clone = value => value == null ? value : JSON.parse(JSON.stringify(value));
  function text(value) {
    if (value == null) return "";
    if (value instanceof Date) return `${value.getUTCFullYear()}/${value.getUTCMonth()+1}/${value.getUTCDate()}`;
    if (value.richText) return value.richText.map(v => v.text).join("");
    if (typeof value === "object") return text(value.result ?? value.text ?? value.formula ?? "");
    return String(value);
  }
  function report(source) { return {source, stages: [], issues: []}; }
  function add(log, code, message, cell="", level="warn", details={}) {
    log.issues.push({source: log.source, code, message, cell, level, ...details});
  }
  function cleanLines(value) {
    return value.replace(/\r\n?/g, "\n").split("\n").map(v=>v.trim()).join("\n").replace(/\n{3,}/g,"\n\n").trim();
  }
  function width(value) {
    // Unicode W/F ranges, including Chinese punctuation and fullwidth ASCII.
    return Array.from(value).reduce((n,ch)=> {
      const c=ch.codePointAt(0);
      const wide = c>=0x1100 && (c<=0x115f || c===0x2329 || c===0x232a ||
        (c>=0x2e80 && c<=0xa4cf && c!==0x303f) || (c>=0xac00 && c<=0xd7a3) ||
        (c>=0xf900 && c<=0xfaff) || (c>=0xfe10 && c<=0xfe19) ||
        (c>=0xfe30 && c<=0xfe6f) || (c>=0xff01 && c<=0xff60) ||
        (c>=0xffe0 && c<=0xffe6) || (c>=0x1f300 && c<=0x1faff) || c>=0x20000);
      return n+(wide?2.2:1);
    },0);
  }
  function colNumber(s) { return Array.from(s).reduce((n,c)=>n*26+c.charCodeAt(0)-64,0); }
  function colLetter(n) { let s=""; for(;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; }
  function range(s) {
    const m=/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(s);
    if(!m) throw new Error(`合并区/区域格式异常：${s}`);
    return [+m[2],colNumber(m[1]),+(m[4]||m[2]),colNumber(m[3]||m[1])];
  }
  const rangeText = ([t,l,b,r])=>`${colLetter(l)}${t}:${colLetter(r)}${b}`;
  function merges(ws, log) {
    const ranges=(ws.model.merges||[]).map(range);
    for(let i=0;i<ranges.length;i++) for(let j=0;j<i;j++) {
      const a=ranges[i], b=ranges[j];
      if(a[0]<=b[2] && b[0]<=a[2] && a[1]<=b[3] && b[1]<=a[3]) {
        add(log,"MERGE_INVALID",`合并区重叠：${rangeText(a)} / ${rangeText(b)}`,"","error");
        throw new Error("合并区异常，停止导出以保护数据");
      }
    }
    return ranges;
  }
  function detach(ws, ranges) {
    const styles=[];
    for(const [t,l,b,r] of ranges) for(let rr=t;rr<=b;rr++) for(let c=l;c<=r;c++) styles.push([rr,c,clone(ws.getCell(rr,c).style)]);
    for(const m of ranges) ws.unMergeCells(...m);
    for(const [r,c,style] of styles) ws.getCell(r,c).style=style;
  }
  function attach(ws, ranges) {
    for(const m of ranges) if(m[0]!==m[2] || m[1]!==m[3]) {
      // ExcelJS otherwise links every slave's style object to the master.
      const styles=[];
      for(let r=m[0];r<=m[2];r++) for(let c=m[1];c<=m[3];c++) styles.push([r,c,clone(ws.getCell(r,c).style)]);
      ws.mergeCells(...m);
      for(const [r,c,s] of styles) ws.getCell(r,c).style=s;
    }
  }
  function columnMap(ws) {
    function mapping(r) {
      const result={};
      for(const [key,words] of Object.entries(KEYS)) {
        result[key]=[];
        for(let c=1;c<=ws.columnCount;c++) if(words.some(w=>text(ws.getCell(r,c).value).replace(/\s+/g,"").includes(w))) result[key].push(c);
      }
      return result;
    }
    let best=-1, header=2, cols;
    for(let r=1;r<=Math.min(ws.rowCount,20);r++) {
      const cm=mapping(r), score=Object.values(cm).filter(v=>v.length).length;
      if(score>best) { best=score; header=r; cols=cm; }
    }
    if(!cols) throw new Error("工作表为空，未找到表头");
    return {header,cols};
  }
  function yellow(cell) {
    const f=cell.fill;
    if(f?.type!=="pattern" || f.pattern!=="solid") return false;
    if([5,13].includes(f.fgColor?.indexed)) return true;
    const rgb=f.fgColor?.argb;
    if(!rgb || rgb.length<6) return false;
    const s=rgb.slice(-6);
    return parseInt(s.slice(0,2),16)>=220 && parseInt(s.slice(2,4),16)>=200 && parseInt(s.slice(4,6),16)<=170;
  }
  function dateFrom(parts) {
    const [y,m,d]=parts.map(Number), v=new Date(Date.UTC(y,m-1,d));
    if(y<100 || v.getUTCFullYear()!==y || v.getUTCMonth()!==m-1 || v.getUTCDate()!==d) throw new Error("日期非法");
    return v;
  }
  function validateTime(value, log, cell, dateValue="") {
    const original=value;
    let s=value.normalize("NFKC").replace(/：/g,":").replace(/[—–~～]|至|到/g,"-");
    const dates=[...s.matchAll(new RegExp(DATE,"g"))], clocks=[...s.matchAll(new RegExp(CLOCK,"g"))];
    let ds=[];
    try {
      ds=dates.map(m=>dateFrom(m.slice(1)));
      if(!dates.length && dateValue) {
        const dm=new RegExp(`^${DATE}$`).exec(text(dateValue).trim());
        if(!dm) throw new Error("日期列不能解析");
        ds=[dateFrom(dm.slice(1))];
      }
    } catch {
      add(log,"DATE_INVALID","日期非法，保留原值",cell,"error",{value:original}); return [original,[]];
    }
    if(!ds.length) add(log,"TIME_DATE_MISSING","时间缺少可确定的日期，保留原值",cell,"error",{value:original});
    if(clocks.length!==2) {
      add(log,"TIME_FORMAT","需两个完整时刻（如08:30-17:30），未自动补造时刻",cell,"error",{value:original}); return [original,ds];
    }
    const times=clocks.map(m=>m.slice(1).map(Number));
    if(times.some(([h,m])=>h>24 || m>59 || (h===24 && m!==0)) || times[0][0]===24) {
      add(log,"TIME_INVALID","时刻超出范围；24:00只允许作结束时间",cell,"error",{value:original}); return [original,ds];
    }
    const remainder=s.replace(new RegExp(CLOCK,"g"),"").replace(new RegExp(DATE,"g"),"");
    if(dates.length>2 || remainder.replace(/[\s-]/g,"") || (dates.length && (dates[0].index>clocks[0].index || (dates.length===2 && !(clocks[0].index+clocks[0][0].length<=dates[1].index && dates[1].index<clocks[1].index))))) {
      add(log,"TIME_FORMAT","时间格式无法唯一解析，保留原值",cell,"error",{value:original}); return [original,ds];
    }
    if(ds.length) {
      const start=+ds[0]+(times[0][0]*60+times[0][1])*60000, end=+ds[ds.length-1]+(times[1][0]*60+times[1][1])*60000;
      if(end<=start) { add(log,"TIME_ORDER","结束时间不晚于开始时间；跨夜请明确结束日期",cell,"error",{value:original}); return [original,ds]; }
    }
    const firstEnd=clocks[0].index+clocks[0][0].length;
    if(!/^\s*-/.test(s.slice(firstEnd,clocks[1].index))) s=s.slice(0,firstEnd)+"-"+s.slice(firstEnd);
    if(s!==original) add(log,"TIME_FIXED","已统一时间符号或补上两个时刻间的“-”",cell,"info",{before:original,after:s});
    return [s,ds];
  }
  function contextRange(value, year, log, label) {
    const full=[...value.matchAll(new RegExp(DATE,"g"))];
    try {
      if(full.length) {
        const ds=full.map(m=>dateFrom(m.slice(1)));
        if(ds[ds.length-1]<ds[0]) throw new Error();
        return [ds[0],ds[ds.length-1]];
      }
      const m=/(\d{1,2})[月.](\d{1,2})日?\s*[-—–~～至]\s*(\d{1,2})[月.](\d{1,2})日?/.exec(value);
      if(m && year) return [dateFrom([year,+m[1],+m[2]]), dateFrom([year+(+m[3]<+m[1]?1:0),+m[3],+m[4]])];
    } catch { add(log,"DATE_CONTEXT_INVALID",`${label}日期非法，保留原值`,"","error"); }
    return null;
  }
  function moveReferences(ws, map) {
    const move=value=>String(value).split(/\s+/).map(part=> {
      const m=range(part), kept=[];
      for(let r=m[0];r<=m[2];r++) if(map(r)!=null) kept.push(map(r));
      return kept.length?rangeText([Math.min(...kept),m[1],Math.max(...kept),m[3]]):"";
    }).filter(Boolean).join(" ");
    if(ws.pageSetup.printArea) ws.pageSetup.printArea=ws.pageSetup.printArea.split("&&").map(move).filter(Boolean).join("&&");
    if(ws.pageSetup.printTitlesRow) {
      const [a,b]=ws.pageSetup.printTitlesRow.split(":").map(Number);
      ws.pageSetup.printTitlesRow=`${map(a)||a}:${map(b)||b}`;
    }
    if(typeof ws.autoFilter==="string") ws.autoFilter=move(ws.autoFilter)||undefined;
    for(const br of ws.rowBreaks||[]) br.id=map(br.id)||map(br.id-1)||1;
    for(const view of ws.views||[]) if(view.state==="frozen" && view.ySplit) view.ySplit=(map(view.ySplit+1)||view.ySplit+1)-1;
    const validations=ws.dataValidations.model, updated={};
    for(const [key,v] of Object.entries(validations)) { const ref=move(key); if(ref) updated[ref]=v; }
    ws.dataValidations.model=updated;
  }
  function preprocess(ws, log) {
    log.stages.push("preprocess");
    const {header,cols}=columnMap(ws);
    for(const key of ["content","time","measures"]) if(!cols[key].length) add(log,key==="measures"?"MEASURES_MISSING":"COLUMN_MISSING",`必需列未找到：${KEYS[key][0]}`,"","error");
    for(const key of ["content","time","date"]) if(cols[key].length>1) add(log,"COLUMN_AMBIGUOUS",`发现多个${KEYS[key][0]}列，逐列校验，不猜列位置`);
    const ranges=merges(ws,log), deleted=[];
    const maxRow=ws.rowCount, maxCol=ws.columnCount;
    for(let r=header+1;r<=maxRow;r++) {
      const cells=Array.from({length:maxCol},(_,i)=>ws.getCell(r,i+1));
      const anchors=cells.filter(c=>!c.isMerged || c.master===c);
      const riskCol=cols.risk.find(c=>text(ws.getCell(header,c).value).includes("作业"))||cols.risk[0];
      const riskValue=riskCol?text(ws.getCell(r,riskCol).value):"";
      const third=/(?:三|3|Ⅲ|III)\s*级|^\s*3\s*$/i.test(riskValue);
      const known=third || /(?:一|二|四|五|六|[12456]|IV|VI|II|V|I)\s*级|^\s*[12456]\s*$/i.test(riskValue);
      const hasYellow=anchors.some(yellow);
      if(hasYellow && (third || !known)) add(log,"YELLOW_RISK_KEPT",third?"三级风险标黄行保留":"标黄行风险级别无法确定，保留待核对",`${ws.name}!A${r}`,third?"info":"warn");
      const example=["例","示例"].includes(text(cells[0].value).trim()) || anchors.some(c=>c.font?.strike || c.value?.richText?.some(v=>v.font?.strike)) || (hasYellow && known && !third);
      const filler=anchors.some(c=>/^填报人(?:及联系方式)?(?:\s*[:：].*)?$/.test(text(c.value).trim()));
      if(example || filler) { deleted.push(r); add(log,"ROW_DELETED",example?"删除示例行":"删除填报人行",`${ws.name}!A${r}`,"info",{original_row:r}); }
    }
    if(deleted.length) {
      const anchors=ranges.map(m=>clone(ws.getCell(m[0],m[1]).value));
      const anchorStyles=ranges.map(m=>clone(ws.getCell(m[0],m[1]).style));
      detach(ws,ranges);
      const map=r=>deleted.includes(r)?null:r-deleted.filter(d=>d<r).length;
      moveReferences(ws,map);
      for(const r of deleted.slice().reverse()) ws.spliceRows(r,1);
      const rebuilt=[];
      ranges.forEach(([t,l,b,r],i)=> {
        const kept=[]; for(let rr=t;rr<=b;rr++) if(map(rr)!=null) kept.push(map(rr));
        if(kept.length) {
          if(deleted.includes(t)) {
            ws.getCell(kept[0],l).value=anchors[i]; ws.getCell(kept[0],l).style=anchorStyles[i];
            add(log,"MERGE_ANCHOR_MOVED","删除合并首行，原正文及样式已移到剩余区域首格",`${ws.name}!${colLetter(l)}${t}`,"info");
          }
          rebuilt.push([kept[0],l,kept[kept.length-1],r]);
        }
      });
      attach(ws,rebuilt);
    }
    const dataDates=[];
    for(let r=header+1;r<=ws.rowCount;r++) {
      if(!Array.from({length:maxCol},(_,c)=>ws.getCell(r,c+1)).some(c=>c.value!=null)) continue;
      let originalRow=r; for(const d of deleted) if(d<=originalRow) originalRow++;
      for(let c=1;c<=maxCol;c++) {
        const cell=ws.getCell(r,c), addr=`${ws.name}!${colLetter(c)}${originalRow}`;
        if((cell.isMerged && cell.master!==cell) || cell.value?.formula || cell.value?.sharedFormula) continue;
        if(typeof cell.value==="string") {
          cell.value=cleanLines(cell.value);
          if(cols.person.includes(c) && cell.value.includes("\n") && /(?:1\d{10}|0\d{2,3}-?\d{7,8})/.test(cell.value)) cell.value=cell.value.replace(/\s*\n\s*/g,"");
        }
        if(cols.date.includes(c) && cell.value!=null) {
          try {
            const dm=new RegExp(`^${DATE}$`).exec(text(cell.value));
            if(!dm) throw new Error();
            dataDates.push([dateFrom(dm.slice(1)),addr]);
          } catch { add(log,"DATE_INVALID","日期列非法或无法解析，保留原值",addr,"error",{value:text(cell.value)}); }
        }
        if(cols.time.includes(c)) {
          const [value,ds]=validateTime(text(cell.value),log,addr,cols.date.length===1?ws.getCell(r,cols.date[0]).value:"");
          if(typeof cell.value==="string") cell.value=value;
          dataDates.push(...ds.map(d=>[d,addr]));
        }
      }
    }
    let title=""; for(let r=1;r<header;r++) for(let c=1;c<=maxCol;c++) { const cell=ws.getCell(r,c); if(!cell.isMerged || cell.master===cell) title+=" "+text(cell.value); }
    title=title.trim();
    const yearMatch=/(20\d{2})(?!\d)/.exec(title+" "+log.source);
    const year=yearMatch?+yearMatch[1]:(dataDates.length?Math.min(...dataDates.map(([d])=>d.getUTCFullYear())):null);
    const contexts=[["标题",contextRange(title,year,log,"标题")],["文件名",contextRange(log.source,year,log,"文件名")]];
    if(contexts.every(([,r])=>r) && contexts[0][1].some((d,i)=>+d!==+contexts[1][1][i])) add(log,"DATE_CONTEXT_MISMATCH","标题与文件名日期范围不一致","","error");
    for(const [label,rng] of contexts) if(rng) for(const [d,addr] of dataDates) if(d<rng[0] || d>rng[1]) add(log,"DATE_MISMATCH",`日期 ${d.toISOString().slice(0,10)} 不在${label}日期范围内`,addr,"error");
    const weeks=[title,log.source].map(v=>/第\s*([0-9一二三四五六七八九十百]+)\s*周/.exec(v));
    if(weeks.every(Boolean) && weeks[0][1]!==weeks[1][1]) add(log,"WEEK_MISMATCH","标题与文件名周次不一致，输出命名采用文件名周次","","error");
    if(!contexts.some(([,r])=>r)) add(log,"DATE_CONTEXT_MISSING","标题/文件名未找到可确定日期范围，未将业务周次强行当作ISO周次");
    let formulas=false; ws.eachRow(row=>row.eachCell(cell=>{if(cell.value?.formula || cell.value?.sharedFormula) formulas=true;}));
    if(formulas || ws.getTables().length || ws.conditionalFormattings?.length) add(log,"REFERENCE_REVIEW","源表含公式、表格或条件格式；插删行后请核对相关引用");
    return {header,cols,title,deletedRows:deleted.length};
  }
  function autoWidth(ws, header, cols, log, preserveWidths=false) {
    log.stages.push("width");
    for(let c=1;c<=ws.columnCount;c++) {
      let longest=0;
      for(let r=header;r<=ws.rowCount;r++) {
        const cell=ws.getCell(r,c);
        if(cell.isMerged && cell.master!==cell) continue;
        for(const line of text(cell.value).split("\n")) longest=Math.max(longest,width(line));
      }
      const target=Math.max(8,Math.min(cols.measures.includes(c)?56:48,longest+2.4));
      ws.getColumn(c).width=cols.measures.includes(c)&&!preserveWidths?target:Math.max(ws.getColumn(c).width||13,target);
    }
  }
  function autoHeight(ws, cols, log, plus=1) {
    if(!Number.isFinite(plus) || plus<0) throw new Error("管控措施行高增量必须为非负有限数字");
    log.stages.push("height");
    const ranges=merges(ws,log), needed=new Map(), vertical=[];
    for(let r=1;r<=ws.rowCount;r++) {
      let height=15;
      for(let c=1;c<=ws.columnCount;c++) {
        const cell=ws.getCell(r,c);
        if(cell.isMerged && cell.master!==cell) continue;
        cell.alignment={...cell.alignment,wrapText:true,vertical:"middle"};
        const value=text(cell.value); if(!value) continue;
        const m=ranges.find(m=>m[0]===r && m[1]===c)||[r,c,r,c];
        let colWidth=0; for(let cc=c;cc<=m[3];cc++) colWidth+=ws.getColumn(cc).width;
        const size=cell.font?.size||11;
        const lines=value.split("\n").reduce((n,s)=>n+Math.max(1,Math.ceil(width(s)*size/11/Math.max(1,colWidth-2.4))),0);
        const h=Math.max(18,lines*size*1.5+6)+((cols.measures.includes(c)||value.includes("管控措施"))?plus:0);
        if(m[2]>r) vertical.push([r,m[2],h]);
        else height=Math.max(height,h);
      }
      needed.set(r,height);
    }
    // Add only the unmet height to a vertical merge's first row.
    for(const [top,bottom,required] of vertical.sort((a,b)=>(a[1]-a[0])-(b[1]-b[0]))) {
      let available=0;for(let r=top;r<=bottom;r++) available+=needed.get(r);
      needed.set(top,needed.get(top)+Math.max(0,required-available));
    }
    for(const [r,height] of needed) ws.getRow(r).height=height;
    return needed;
  }
  function splitOverflow(ws, needed, log) {
    log.stages.push("insert");
    const ranges=merges(ws,log), maxCol=ws.columnCount;
    detach(ws,ranges);
    let inserted=0;
    for(const [r,required] of [...needed].reverse()) {
      if(required<=MAX_HEIGHT) continue;
      const count=Math.ceil(required/MAX_HEIGHT)-1;
      if(ws.rowCount+count>1048576) throw new Error("插行将超出Excel最大行数");
      const styles=Array.from({length:maxCol},(_,c)=>clone(ws.getCell(r,c+1).style));
      const covered=new Set();
      for(const [t,l,b,rr] of ranges) if(t<=r && r<=b) for(let c=l;c<=rr;c++) covered.add(c);
      moveReferences(ws,old=>old>r?old+count:old);
      // Required order: insert, copy styles, finally rebuild merges after all row moves.
      for(let i=0;i<count;i++) ws.spliceRows(r+1,0,[]);
      for(let nr=r+1;nr<=r+count;nr++) for(let c=1;c<=maxCol;c++) ws.getCell(nr,c).style=clone(styles[c-1]);
      for(let nr=r;nr<=r+count;nr++) ws.getRow(nr).height=required/(count+1);
      for(const m of ranges) { if(m[0]>r) m[0]+=count; if(m[2]>=r) m[2]+=count; }
      for(let c=1;c<=maxCol;c++) if(!covered.has(c)) ranges.push([r,c,r+count,c]);
      inserted+=count;
      add(log,"ROW_SPLIT",`整行需求${required.toFixed(1)}磅，分摊至${count+1}行`,`${ws.name}!A${r}`,"info",{inserted:count});
    }
    attach(ws,ranges);
    for(let r=1;r<=ws.rowCount;r++) for(let c=1;c<=maxCol;c++) {
      const cell=ws.getCell(r,c); cell.alignment={...cell.alignment,wrapText:true,vertical:"middle"};
    }
    return inserted;
  }
  function outputName(filename, title="") {
    return filename.split(/[\\/]/).pop();
  }
  function detailName(filename, title="") {
    const stem=filename.replace(/\.xlsx$/i,"");
    const m=/第\s*([0-9一二三四五六七八九十百]+)\s*周/.exec(stem)||/第\s*([0-9一二三四五六七八九十百]+)\s*周/.exec(title);
    return m?`（第${m[1]}周）（周计划明细）.xlsx`:"周计划明细.xlsx";
  }
  function setupSourcePrint(ws, header, cols, log) {
    const originalScale=ws.pageSetup.scale;
    let printCol=ws.columnCount;
    if(ws.pageSetup.printArea) {
      const right=ws.pageSetup.printArea.split('&&').map(s=>range(s)[3]).reduce((n,c)=>Math.max(n,c),0);
      if(right>=Math.max(...cols.measures)) printCol=Math.min(printCol,right);
    }
    ws.pageSetup={...ws.pageSetup,orientation:"landscape",paperSize:8,fitToPage:true,fitToWidth:1,fitToHeight:0,
      margins:{left:0.3,right:0.3,top:0.35,bottom:0.35,header:0.15,footer:0.15},
      printArea:`A1:${colLetter(printCol)}${ws.rowCount}`,printTitlesRow:`1:${header}`};
    delete ws.pageSetup.scale;
    ws.rowBreaks=[];
    let widthPoints=0;
    for(let c=1;c<=printCol;c++) widthPoints+=(ws.getColumn(c).width*7+5)*0.75;
    let scale=Math.min(1,Math.max(0.1,(1190.55-43.2)/widthPoints));
    if(originalScale>=10 && originalScale<=400 && widthPoints*originalScale/100<=1147.35) {
      ws.pageSetup.scale=originalScale;ws.pageSetup.fitToPage=false;scale=originalScale/100;
    }
    let headerHeight=0;for(let r=1;r<=header;r++) headerHeight+=ws.getRow(r).height||15;
    const budget=(841.89-50.4-24)/scale-headerHeight;
    const ranges=merges(ws,log);
    const totalHeight=(a,b)=>{let h=0;for(let r=a;r<=b;r++) h+=ws.getRow(r).height||15;return h;};
    const pageBreak=r=>ws.getRow(r).addPageBreak(0,ws.columnCount-1);
    let used=0,r=header+1;
    while(r<=ws.rowCount) {
      const end=Math.max(r,...ranges.filter(m=>m[0]<=r && r<=m[2] && m[1]<=2 && m[3]<=2).map(m=>m[2]));
      const total=totalHeight(r,end);
      if(total<=budget) {
        if(used && used+total>budget) {pageBreak(r-1);used=0;}
        used+=total;
      } else {
        if(used) {pageBreak(r-1);used=0;}
        let rr=r;
        while(rr<=end) {
          const atomicEnd=Math.min(end,Math.max(rr,...ranges.filter(m=>m[0]<=rr && rr<=m[2] && cols.measures.some(c=>m[1]<=c && c<=m[3])).map(m=>m[2])));
          const height=totalHeight(rr,atomicEnd);
          if(used && used+height>budget) {pageBreak(rr-1);used=0;}
          if(height>budget) add(log,"PRINT_BLOCK_OVERSIZE","单个管控措施合并块超过A3一页可用高度，请核对打印预览；未截断或删除正文",`${ws.name}!A${rr}`);
          used+=height;rr=atomicEnd+1;
        }
      }
      r=end+1;
    }
    const mode=ws.pageSetup.scale?`保留原打印缩放${originalScale}%`:'宽度一页、高度不限';
    add(log,"A3_PRINT",`A3横向、${mode}；已按作业/管控措施合并块设置${ws.rowBreaks.length}处分分页，原字体和加粗保留`,"","info");
  }
  function isSource(filename) {
    return /\.xlsx$/i.test(filename) && !/(?:^|[\\/])~\$|明细|处理后|(?:^|[\\/])work(?:[\\/]|$)|源文件备份/i.test(filename);
  }
  async function trimInflatedCopy(buffer, log) {
    const zip=await ExcelJS.SourceZip.loadAsync(buffer);
    const book=await zip.file('xl/workbook.xml').async('string');
    const first=/<sheet\b[^>]*\br:id="([^"]+)"/.exec(book);
    if(!first) return buffer;
    const rels=await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rel=[...rels.matchAll(/<Relationship\b[^>]*\/>/g)].find(m=>new RegExp(`\\bId="${first[1]}"`).test(m[0]));
    if(!rel) return buffer;
    const target=/\bTarget="([^"]+)"/.exec(rel[0])[1];
    const name=target.startsWith('/')?target.slice(1):'xl/'+target;
    let xml=await zip.file(name).async('string');
    const cellRE=/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
    const cells=[...xml.matchAll(cellRE)];
    const col=s=>colNumber(/\br="([A-Z]+)\d+"/.exec(s)[1]);
    const populated=cells.filter(m=>/<(?:v|is|f)(?:\s|>)/.test(m[0])).map(m=>col(m[0]));
    if(!populated.length) return buffer;
    let bound=populated.reduce((n,c)=>Math.max(n,c),0);
    const ranges=[...xml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"[^>]*\/>/g)].map(m=>m[1]);
    for(const ref of ranges) {const m=range(ref);if(m[1]<=bound) bound=Math.max(bound,m[3]);}
    const actual=cells.reduce((n,m)=>Math.max(n,col(m[0])),0);
    if(actual<=bound+256) return buffer;
    xml=xml.replace(cellRE,m=>col(m)<=bound?m:'');
    const kept=ranges.filter(ref=>range(ref)[1]<=bound);
    xml=xml.replace(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/,()=>`<mergeCells count="${kept.length}">${kept.map(ref=>`<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`);
    xml=xml.replace(/<col\b[^>]*\/>/g,m=>+(/\bmin="(\d+)"/.exec(m)[1])>bound?'':m.replace(/\bmax="(\d+)"/,(_,hi)=>`max="${Math.min(bound,+hi)}"`));
    xml=xml.replace(/<dimension\b[^>]*\/>/,'');
    zip.file(name,xml);
    add(log,'EMPTY_FORMAT_TRIMMED',`工作副本移除${colLetter(bound+1)}列以外的远端空白格式及${ranges.length-kept.length}个空合并；正文未改`,'','info');
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  }
  async function load(buffer, log) {
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await trimInflatedCopy(buffer,log));
    if(!wb.worksheets.length) throw new Error("工作簿没有工作表");
    // ExcelJS loses localSheetId for this Excel-owned name and spliceRows fragments it.
    // Keep worksheet.autoFilter, which is the authoritative filter range. A global reserved
    // _FilterDatabase name makes desktop Excel reject the entire otherwise readable file.
    wb.definedNames.model=wb.definedNames.model.filter(n=>n.name!=="_xlnm._FilterDatabase");
    return wb;
  }
  async function preprocessBuffer(buffer, filename="表1.xlsx") {
    const log=report(filename), wb=await load(buffer,log);
    const meta=preprocess(wb.worksheets[0],log);
    return {buffer:await wb.xlsx.writeBuffer(),...meta,report:log,actions:log.issues.map(i=>i.message)};
  }
  async function processBuffer(buffer, filename="表1.xlsx", plus=1, preserveWidths=true) {
    const log=report(filename);
    try {
      const wb=await load(buffer,log), ws=wb.worksheets[0];
      const meta=preprocess(ws,log);
      autoWidth(ws,meta.header,meta.cols,log,preserveWidths);
      const needed=autoHeight(ws,meta.cols,log,plus);
      const inserted=splitOverflow(ws,needed,log);
      if(wb.worksheets.length>1) add(log,"OTHER_SHEETS","仅处理第一张工作表，其余工作表保留","","info");
      setupSourcePrint(ws,meta.header,meta.cols,log);
      const out=await wb.xlsx.writeBuffer(); log.stages.push("output");
      return {buffer:out,report:log,inserted,adjusted:needed.size,filename:outputName(filename,meta.title),...meta};
    } catch(e) {
      add(log,/merge|合并/i.test(e.message)?"MERGE_INVALID":"OUTPUT_FAILED",`处理/输出失败：${e.message}`,"","error");
      e.report=log; throw e;
    }
  }
  // Backups are keyed by original filename, deliberately not by the changed content hash.
  const memoryBackups=new Map(), work=new Map();
  let dbPromise;
  function openBackups() {
    if(!dbPromise) dbPromise=new Promise((resolve,reject)=> {
      if(typeof indexedDB==="undefined") {reject(new Error("浏览器不支持持久备份"));return;}
      const request=indexedDB.open("weekly-plan-source-backups-v1",1);
      request.onupgradeneeded=()=>request.result.createObjectStore("originals");
      request.onsuccess=()=>resolve(request.result);
      request.onerror=()=>reject(request.error);
    });
    return dbPromise;
  }
  async function backupOnce(filename, bytes, log) {
    try {
      const db=await openBackups();
      await new Promise((resolve,reject)=> {
        const tx=db.transaction("originals","readwrite"), store=tx.objectStore("originals"), get=store.get(filename);
        get.onsuccess=()=>{if(get.result===undefined) store.add(bytes.slice(0),filename);else add(log,"BACKUP_EXISTS","备份已存在，跳过且不覆盖","","info");};
        tx.oncomplete=resolve; tx.onerror=()=>reject(tx.error); tx.onabort=()=>reject(tx.error||new Error("备份中断"));
      });
    } catch(e) {
      if(!memoryBackups.has(filename)) memoryBackups.set(filename,bytes.slice(0));
      add(log,"BACKUP_MEMORY_ONLY","浏览器持久备份不可用，仅保留本次会话备份；请下载源文件备份ZIP。"+e.message);
    }
  }
  async function getBackup(filename) {
    try {
      const db=await openBackups();
      const value=await new Promise((resolve,reject)=> {
        const req=db.transaction("originals").objectStore("originals").get(filename);
        req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
      });
      if(value!==undefined) return value;
    } catch { /* fallback to this session's explicitly reported in-memory backup */ }
    const value=memoryBackups.get(filename);
    if(value===undefined) throw new Error("未找到源文件备份，请重新导入源文件");
    return value;
  }
  async function processFile(file, plus=1, preserveWidths=true) {
    if(!isSource(file.name)) throw new Error("请导入原始.xlsx源表；已排除临时文件、明细、处理后源表和work目录");
    const bytes=await file.arrayBuffer(), log=report(file.name);
    await backupOnce(file.name,bytes,log); log.stages.push("backup_work");
    const token=Symbol(file.name); work.set(token,bytes.slice(0));
    let out;
    try {
      out=await processBuffer(work.get(token),file.name,plus,preserveWidths);
      out.report.issues.unshift(...log.issues); out.report.stages.unshift(...log.stages);
      return out;
    } catch(e) {
      if(e.report) {e.report.issues.unshift(...log.issues);e.report.stages.unshift(...log.stages);e.report.stages.push("work_cleanup");}
      throw e;
    } finally { work.delete(token); if(out) out.report.stages.push("work_cleanup"); }
  }
  // Minimal ZIP STORE writer: backup bytes are copied verbatim under 源文件备份/.
  async function backupZip(filename) {
    const bytes=new Uint8Array(await getBackup(filename)), name=new TextEncoder().encode("源文件备份/"+filename);
    let crc=0xffffffff;
    for(const b of bytes) {crc^=b;for(let i=0;i<8;i++) crc=(crc>>>1)^((crc&1)?0xedb88320:0);}
    crc=(crc^0xffffffff)>>>0;
    const result=new Uint8Array(30+name.length+bytes.length+46+name.length+22), v=new DataView(result.buffer);
    const u16=(p,n)=>v.setUint16(p,n,true), u32=(p,n)=>v.setUint32(p,n,true);
    u32(0,0x04034b50);u16(4,20);u16(6,0x800);u16(12,33);u32(14,crc);u32(18,bytes.length);u32(22,bytes.length);u16(26,name.length);
    result.set(name,30);result.set(bytes,30+name.length);
    const p=30+name.length+bytes.length;
    u32(p,0x02014b50);u16(p+4,20);u16(p+6,20);u16(p+8,0x800);u16(p+14,33);u32(p+16,crc);u32(p+20,bytes.length);u32(p+24,bytes.length);u16(p+28,name.length);result.set(name,p+46);
    const e=p+46+name.length;
    u32(e,0x06054b50);u16(e+8,1);u16(e+10,1);u32(e+12,46+name.length);u32(e+16,p);
    return result;
  }
  async function outputZip(files) {
    const zip=new ExcelJS.SourceZip();
    for(const file of files) {
      if(!file.name || /[\\/]/.test(file.name) || file.name==='.' || file.name==='..') throw new Error('输出文件名无效');
      zip.file('（处理后）/'+file.name,file.data);
    }
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  }
  return {processBuffer,processFile,preprocessBuffer,outputName,detailName,outputZip,isSource,backupZip,getBackup,backupOnce,
    work,report,preprocess,autoWidth,autoHeight,splitOverflow,validateTime,width,columnMap};
}
if (typeof module !== "undefined" && module.exports) module.exports = {createSourceProcessor};
