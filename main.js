const ExcelJS = require('exceljs');

// ── カラー定義 ──────────────────────────────────────
const C = {
  navyBg: '1A3557', navyFg: 'FFFFFFFF',
  blueBg: '2E75B6', blueFg: 'FFFFFFFF',
  s1:     'FFE8F0FB', s2:   'FFFFFFFF',
  totBg:  'FFD6E4F7', totFg:'FF1A3557',
  wBg:    'FFFFF2CC', wFg:  'FF7F4F00',
  bdr:    'FFBDD0E8',
};

// ── スタイルヘルパー ────────────────────────────────
const fill  = argb => ({ type:'pattern', pattern:'solid', fgColor:{argb} });
const fnt   = (bold, argb, size=10) => ({ name:'Meiryo UI', bold, color:{argb}, size });
const bord  = (argb=C.bdr) => {
  const s = { style:'thin', color:{argb} };
  return { top:s, bottom:s, left:s, right:s };
};

const ST = {
  hdrL: { fill:fill('FF'+C.navyBg), font:fnt(true,'FFFFFFFF'), border:bord('FF4A6080'), alignment:{horizontal:'left',  vertical:'middle'} },
  hdrR: { fill:fill('FF'+C.navyBg), font:fnt(true,'FFFFFFFF'), border:bord('FF4A6080'), alignment:{horizontal:'right', vertical:'middle'} },
  subL: { fill:fill('FF'+C.blueBg), font:fnt(true,'FFFFFFFF'), border:bord('FF5080A0'), alignment:{horizontal:'left',  vertical:'middle'} },
  titl: { fill:fill('FFFFFFFF'),    font:fnt(true,'FF'+C.navyBg, 13),                       alignment:{horizontal:'left',  vertical:'middle'} },
  note: { fill:fill('FFFFFFFF'),    font:{ name:'Meiryo UI', size:9, color:{argb:'FF888888'}, italic:true }, alignment:{horizontal:'left', vertical:'middle'} },
  totL: { fill:fill(C.totBg),  font:fnt(true,C.totFg),  border:bord(C.bdr), alignment:{horizontal:'left',  vertical:'middle'}, numFmt:'@'     },
  totN: { fill:fill(C.totBg),  font:fnt(true,C.totFg),  border:bord(C.bdr), alignment:{horizontal:'right', vertical:'middle'}, numFmt:'#,##0' },
  totP: { fill:fill(C.totBg),  font:fnt(true,C.totFg),  border:bord(C.bdr), alignment:{horizontal:'right', vertical:'middle'}, numFmt:'0.0%'  },
  dL:  e => ({ fill:fill(e?C.s1:C.s2), font:fnt(false,'FF333333'), border:bord(C.bdr), alignment:{horizontal:'left',  vertical:'middle'} }),
  dN:  e => ({ fill:fill(e?C.s1:C.s2), font:fnt(false,'FF333333'), border:bord(C.bdr), alignment:{horizontal:'right', vertical:'middle'}, numFmt:'#,##0' }),
  dP:  e => ({ fill:fill(e?C.s1:C.s2), font:fnt(false,'FF333333'), border:bord(C.bdr), alignment:{horizontal:'right', vertical:'middle'}, numFmt:'0.0%'  }),
  wL:  { fill:fill(C.wBg), font:fnt(false,C.wFg), border:bord(C.bdr), alignment:{horizontal:'left',  vertical:'middle'} },
  wN:  { fill:fill(C.wBg), font:fnt(false,C.wFg), border:bord(C.bdr), alignment:{horizontal:'right', vertical:'middle'}, numFmt:'#,##0' },
  wP:  { fill:fill(C.wBg), font:fnt(false,C.wFg), border:bord(C.bdr), alignment:{horizontal:'right', vertical:'middle'}, numFmt:'0.0%'  },
};

// スタイルをセルに適用
function applyStyle(cell, st) {
  if (st.fill)      cell.fill      = st.fill;
  if (st.font)      cell.font      = st.font;
  if (st.border)    cell.border    = st.border;
  if (st.alignment) cell.alignment = st.alignment;
  if (st.numFmt)    cell.numFmt    = st.numFmt;
}

// ── ユーティリティ ──────────────────────────────────
const toNum = v => { const n = parseFloat(String(v).replace(/,/g,'')); return isNaN(n)?0:n; };
const isLow = (u,a) => u>0 && a/u<0.10;

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter(Boolean);
  const headers = splitLine(lines[0]).map(h => h.trim().replace(/^\uFEFF/,''));
  const rows = lines.slice(1).map(l => {
    const vals = splitLine(l);
    const row = {};
    headers.forEach((h,i) => { row[h] = (vals[i]??'').trim(); });
    return row;
  });
  return { headers, rows };
}
function splitLine(line) {
  const r=[]; let cur=''; let inQ=false;
  for(const c of line){
    if(c==='"') inQ=!inQ;
    else if(c===','&&!inQ){r.push(cur);cur='';}
    else cur+=c;
  }
  r.push(cur); return r;
}

// ── シートヘルパー ──────────────────────────────────
function writeTitle(ws, title, note, ncols) {
  const r1 = ws.addRow([title]);
  r1.height = 28;
  const tc = r1.getCell(1);
  tc.value = title; applyStyle(tc, ST.titl);
  if(ncols>1) ws.mergeCells(r1.number, 1, r1.number, ncols);

  const r2 = ws.addRow([note]);
  r2.height = 16;
  const nc = r2.getCell(1);
  nc.value = note; applyStyle(nc, ST.note);
  if(ncols>1) ws.mergeCells(r2.number, 1, r2.number, ncols);

  ws.addRow([]);
}

function writeSubHeader(ws, label, ncols) {
  const row = ws.addRow([label]);
  row.height = 20;
  for(let c=1;c<=ncols;c++) applyStyle(row.getCell(c), ST.subL);
  if(ncols>1) ws.mergeCells(row.number,1,row.number,ncols);
}

function writeHeader(ws, labels) {
  const row = ws.addRow(labels);
  row.height = 20;
  labels.forEach((_,i) => applyStyle(row.getCell(i+1), i===0?ST.hdrL:ST.hdrR));
}

function setColWidths(ws, widths) {
  widths.forEach((w,i) => { ws.getColumn(i+1).width = w; });
}

// ── 集計ヘルパー ────────────────────────────────────
function jsAgg(rows, key) {
  const m={};
  for(const r of rows){
    const k=r[key]||'（未設定）';
    if(!m[k])m[k]={u:0,a:0};
    m[k].u+=toNum(r['売上金額']);
    m[k].a+=toNum(r['粗利金額']);
  }
  return Object.entries(m).map(([k,v])=>({name:k,...v})).sort((x,y)=>y.a-x.a);
}

// ── SUMIF数式（ローデータ参照） ─────────────────────
const RAW = '⑧ローデータ';
function sf(colName, val, sumCol, dr, colIdx) {
  const cc = colLetter(colIdx[colName]);
  const sc = colLetter(colIdx[sumCol]);
  return { formula: `SUMIF('${RAW}'!${cc}$2:${cc}$${dr+1},"${val}",'${RAW}'!${sc}$2:${sc}$${dr+1})` };
}
function total(sumCol, dr, colIdx) {
  const sc = colLetter(colIdx[sumCol]);
  return { formula: `SUM('${RAW}'!${sc}$2:${sc}$${dr+1})` };
}
function cntif(colName, val, dr, colIdx) {
  const cc = colLetter(colIdx[colName]);
  return { formula: `COUNTIF('${RAW}'!${cc}$2:${cc}$${dr+1},"${val}")` };
}
function colLetter(idx) {
  let s=''; idx++;
  while(idx>0){idx--;s=String.fromCharCode(65+idx%26)+s;idx=Math.floor(idx/26);}
  return s;
}

// ════════════════════════════════════════════════════
// メイン：ExcelJSでワークブック構築
// ════════════════════════════════════════════════════
async function buildWorkbook(csvText) {
  const { headers, rows } = parseCSV(csvText);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ビルマテル粗利分析ツール';

  const DR = rows.length;
  const colIdx = {};
  headers.forEach((h,i) => { colIdx[h]=i; });

  const uniq = key => [...new Set(rows.map(r=>r[key]||'（未設定）'))].sort();
  const customers = uniq('得意先名');
  const staffs    = uniq('担当者名');
  const products  = uniq('売上分類名');
  const sites     = uniq('現場名');

  // 月次はJS集計
  const monthMap={};
  for(const r of rows){
    const d=(r['売上日']||'').trim().replace(/\//g,'-');
    const m=d.slice(0,7);
    if(!m||m.length<7) continue;
    if(!monthMap[m])monthMap[m]={u:0,a:0};
    monthMap[m].u+=toNum(r['売上金額']);
    monthMap[m].a+=toNum(r['粗利金額']);
  }
  const months=Object.keys(monthMap).sort();

  const totU = rows.reduce((s,r)=>s+toNum(r['売上金額']),0);
  const totA = rows.reduce((s,r)=>s+toNum(r['粗利金額']),0);

  // ── ① サマリー ──────────────────────────────────
  {
    const ws = wb.addWorksheet('①サマリー');
    setColWidths(ws,[28,16,16,10,12,12]);
    writeTitle(ws,'粗利分析レポート　サマリー','※ 粗利率 = 粗利金額 ÷ 売上金額（CAULの粗利率列は不使用）',6);

    writeSubHeader(ws,'■ 全体集計',3);
    const kpis=[
      ['売上金額合計', total('売上金額',DR,colIdx), '#,##0', '円'],
      ['粗利金額合計', total('粗利金額',DR,colIdx), '#,##0', '円'],
      ['粗利率（全体）', {formula:`IFERROR(${total('粗利金額',DR,colIdx).formula}/${total('売上金額',DR,colIdx).formula},0)`}, '0.0%', ''],
      ['対象件数', rows.length, '#,##0', '件'],
      ['得意先数', customers.length, '#,##0', '社'],
      ['担当者数', staffs.length, '#,##0', '名'],
    ];
    kpis.forEach(([lb,v,fmt,unit],i)=>{
      const e=i%2===0;
      const row=ws.addRow([lb,typeof v==='object'?v:v,unit]);
      row.height=18;
      applyStyle(row.getCell(1),ST.dL(e));
      const c2=row.getCell(2); c2.value=v; applyStyle(c2,fmt==='0.0%'?ST.dP(e):ST.dN(e)); c2.numFmt=fmt;
      applyStyle(row.getCell(3),ST.dL(e));
    });

    const custSorted=jsAgg(rows,'得意先名');
    const writeRank=(title,list)=>{
      ws.addRow([]);
      writeSubHeader(ws,title,6);
      writeHeader(ws,['得意先名','売上金額','粗利金額','粗利率','売上構成比','粗利構成比']);
      list.forEach((kd,i)=>{
        const e=i%2===0, low=isLow(kd.u,kd.a);
        const uF=sf('得意先名',kd.name,'売上金額',DR,colIdx);
        const aF=sf('得意先名',kd.name,'粗利金額',DR,colIdx);
        const row=ws.addRow([]); row.height=18;
        const c1=row.getCell(1); c1.value=kd.name; applyStyle(c1,low?ST.wL:ST.dL(e));
        const c2=row.getCell(2); c2.value=uF;       applyStyle(c2,low?ST.wN:ST.dN(e));
        const c3=row.getCell(3); c3.value=aF;       applyStyle(c3,low?ST.wN:ST.dN(e));
        const rn=row.number;
        const c4=row.getCell(4); c4.value={formula:`IFERROR(C${rn}/B${rn},0)`}; applyStyle(c4,low?ST.wP:ST.dP(e));
        const c5=row.getCell(5); c5.value={formula:`IFERROR(B${rn}/${total('売上金額',DR,colIdx).formula},0)`}; applyStyle(c5,low?ST.wP:ST.dP(e));
        const c6=row.getCell(6); c6.value={formula:`IFERROR(C${rn}/${total('粗利金額',DR,colIdx).formula},0)`}; applyStyle(c6,low?ST.wP:ST.dP(e));
      });
    };
    writeRank('■ 粗利金額 TOP5 得意先', custSorted.slice(0,5));
    writeRank('■ 粗利金額 BOTTOM5 得意先（低粗利注意）', [...custSorted].reverse().slice(0,5));
    ws.views=[{state:'normal'}];
  }

  // ── ②〜⑤ 軸別（共通） ──────────────────────────
  const makeAxis=(sname,title,colName)=>{
    const ws=wb.addWorksheet(sname);
    setColWidths(ws,[30,16,16,10,12,12,10]);
    writeTitle(ws,`粗利分析　${title}`,'※ 粗利率 = 粗利金額 ÷ 売上金額（黄色ハイライト = 粗利率10%未満）',7);
    writeHeader(ws,[title,'売上金額','粗利金額','粗利率','売上構成比','粗利構成比','件数']);
    jsAgg(rows,colName).forEach((kd,i)=>{
      const e=i%2===0, low=isLow(kd.u,kd.a);
      const uF=sf(colName,kd.name,'売上金額',DR,colIdx);
      const aF=sf(colName,kd.name,'粗利金額',DR,colIdx);
      const row=ws.addRow([]); row.height=18;
      const c1=row.getCell(1); c1.value=kd.name; applyStyle(c1,low?ST.wL:ST.dL(e));
      const c2=row.getCell(2); c2.value=uF;       applyStyle(c2,low?ST.wN:ST.dN(e));
      const c3=row.getCell(3); c3.value=aF;       applyStyle(c3,low?ST.wN:ST.dN(e));
      const rn=row.number;
      const c4=row.getCell(4); c4.value={formula:`IFERROR(C${rn}/B${rn},0)`};                                  applyStyle(c4,low?ST.wP:ST.dP(e));
      const c5=row.getCell(5); c5.value={formula:`IFERROR(B${rn}/${total('売上金額',DR,colIdx).formula},0)`};  applyStyle(c5,low?ST.wP:ST.dP(e));
      const c6=row.getCell(6); c6.value={formula:`IFERROR(C${rn}/${total('粗利金額',DR,colIdx).formula},0)`}; applyStyle(c6,low?ST.wP:ST.dP(e));
      const c7=row.getCell(7); c7.value=cntif(colName,kd.name,DR,colIdx); applyStyle(c7,ST.dN(e));
    });
    // 合計行
    const tot=ws.addRow([]); tot.height=18;
    const lastData=ws.lastRow.number;
    const hdrRow=4; // タイトル3行+空行=4行目がヘッダ
    const c1=tot.getCell(1); c1.value='合　計'; applyStyle(c1,ST.totL);
    const c2=tot.getCell(2); c2.value=total('売上金額',DR,colIdx); applyStyle(c2,ST.totN);
    const c3=tot.getCell(3); c3.value=total('粗利金額',DR,colIdx); applyStyle(c3,ST.totN);
    const rn=tot.number;
    const c4=tot.getCell(4); c4.value={formula:`IFERROR(C${rn}/B${rn},0)`}; applyStyle(c4,ST.totP);
    const c5=tot.getCell(5); c5.value='100.0%'; applyStyle(c5,ST.totL);
    const c6=tot.getCell(6); c6.value='100.0%'; applyStyle(c6,ST.totL);
    const c7=tot.getCell(7); c7.value={formula:`COUNTA('${RAW}'!${colLetter(colIdx[colName])}$2:${colLetter(colIdx[colName])}$${DR+1})`}; applyStyle(c7,ST.totN);
  };
  makeAxis('②得意先別',  '得意先名',  '得意先名');
  makeAxis('③担当者別',  '担当者名',  '担当者名');
  makeAxis('④商品分類別','売上分類名','売上分類名');
  makeAxis('⑤現場別',   '現場名',    '現場名');

  // ── ⑥ 月次推移 ──────────────────────────────────
  {
    const ws=wb.addWorksheet('⑥月次推移');
    setColWidths(ws,[14,16,16,10,16]);
    writeTitle(ws,'粗利分析　月次推移','※ 粗利率 = 粗利金額 ÷ 売上金額',5);
    writeHeader(ws,['年月','売上金額','粗利金額','粗利率','前月差（粗利率）']);
    let prevRateRowNum=null;
    months.forEach((m,i)=>{
      const e=i%2===0;
      const {u,a}=monthMap[m];
      const row=ws.addRow([]); row.height=18;
      row.getCell(1).value=m;   applyStyle(row.getCell(1),ST.dL(e));
      row.getCell(2).value=u;   applyStyle(row.getCell(2),ST.dN(e));
      row.getCell(3).value=a;   applyStyle(row.getCell(3),ST.dN(e));
      const rn=row.number;
      row.getCell(4).value={formula:`IFERROR(C${rn}/B${rn},0)`}; applyStyle(row.getCell(4),ST.dP(e));
      if(i===0){
        row.getCell(5).value='—'; applyStyle(row.getCell(5),ST.dL(e));
      } else {
        row.getCell(5).value={formula:`IFERROR(D${rn}-D${prevRateRowNum},0)`}; applyStyle(row.getCell(5),ST.dP(e));
      }
      prevRateRowNum=rn;
    });
    const tot=ws.addRow([]); tot.height=18;
    const rn=tot.number;
    const hdr=5;
    tot.getCell(1).value='合　計'; applyStyle(tot.getCell(1),ST.totL);
    tot.getCell(2).value={formula:`SUM(B${hdr}:B${rn-1})`}; applyStyle(tot.getCell(2),ST.totN);
    tot.getCell(3).value={formula:`SUM(C${hdr}:C${rn-1})`}; applyStyle(tot.getCell(3),ST.totN);
    tot.getCell(4).value={formula:`IFERROR(C${rn}/B${rn},0)`}; applyStyle(tot.getCell(4),ST.totP);
    tot.getCell(5).value=''; applyStyle(tot.getCell(5),ST.totL);
  }

  // ── ⑦ 得意先×担当者 ─────────────────────────────
  {
    const ws=wb.addWorksheet('⑦得意先×担当者');
    setColWidths(ws,[30,14,16,16,10,10]);
    writeTitle(ws,'粗利分析　得意先 × 担当者 クロス','※ 粗利率 = 粗利金額 ÷ 売上金額（黄色ハイライト = 粗利率10%未満）',6);
    writeHeader(ws,['得意先名','担当者名','売上金額','粗利金額','粗利率','件数']);
    const cross=[];
    for(const cu of customers) for(const st of staffs){
      let u=0,a=0,cnt=0;
      for(const r of rows){
        if((r['得意先名']||'（未設定）')===cu&&(r['担当者名']||'（未設定）')===st){
          u+=toNum(r['売上金額']); a+=toNum(r['粗利金額']); cnt++;
        }
      }
      if(u>0) cross.push({cu,st,u,a,cnt});
    }
    cross.sort((x,y)=>y.a-x.a).forEach((cd,i)=>{
      const e=i%2===0, low=isLow(cd.u,cd.a);
      const row=ws.addRow([]); row.height=18;
      row.getCell(1).value=cd.cu;  applyStyle(row.getCell(1),low?ST.wL:ST.dL(e));
      row.getCell(2).value=cd.st;  applyStyle(row.getCell(2),low?ST.wL:ST.dL(e));
      row.getCell(3).value=cd.u;   applyStyle(row.getCell(3),low?ST.wN:ST.dN(e));
      row.getCell(4).value=cd.a;   applyStyle(row.getCell(4),low?ST.wN:ST.dN(e));
      const rn=row.number;
      row.getCell(5).value={formula:`IFERROR(D${rn}/C${rn},0)`}; applyStyle(row.getCell(5),low?ST.wP:ST.dP(e));
      row.getCell(6).value=cd.cnt; applyStyle(row.getCell(6),ST.dN(e));
    });
  }

  // ── ⑧ ローデータ ────────────────────────────────
  {
    const ws=wb.addWorksheet('⑧ローデータ');
    const numCols=new Set(['数量','売上単価','売上金額','原価単価','原価金額','粗利金額','粗利率','仕入NO','売上NO','売上NO行']);
    const estW=str=>{let w=0;for(const c of String(str??''))w+=c.charCodeAt(0)>127?2.2:1.1;return Math.max(8,Math.min(w+2,30));};

    // ヘッダ
    const hrow=ws.addRow(headers); hrow.height=20;
    headers.forEach((_,i)=>applyStyle(hrow.getCell(i+1),ST.hdrL));
    ws.getRow(1).freeze=true;
    ws.views=[{state:'frozenRows',ySplit:1}];

    // 列幅
    headers.forEach((h,i)=>{
      const maxW=Math.max(estW(h),...rows.slice(0,300).map(r=>estW(r[h])));
      ws.getColumn(i+1).width=Math.min(maxW,30);
    });

    // データ
    rows.forEach((r,ri)=>{
      const e=ri%2===0;
      const vals=headers.map(h=>r[h]??'');
      const row=ws.addRow(vals); row.height=18;
      headers.forEach((h,ci)=>{
        const cell=row.getCell(ci+1);
        const raw=r[h]??'';
        const n=parseFloat(String(raw).replace(/,/g,''));
        const isN=numCols.has(h)&&!isNaN(n)&&raw!=='';
        if(isN){
          cell.value=n;
          const fmt=h==='粗利率'?'0.0%':'#,##0';
          applyStyle(cell,{...ST.dN(e),numFmt:fmt});
        } else {
          applyStyle(cell,ST.dL(e));
        }
      });
    });
  }

  return wb;
}

// ── ブラウザ向けエクスポート ────────────────────────
window.buildAndDownload = async function(csvText) {
  const wb = await buildWorkbook(csvText);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download='粗利分析レポート.xlsx'; a.click();
  URL.revokeObjectURL(url);
};
