const express = require('express');
const multer = require('multer');
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const crypto = require('crypto');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const PUBLIC = path.join(ROOT, 'public');
const DATA = path.join(ROOT, 'data');
const BACKUPS = path.join(ROOT, 'backups');
const TMP = path.join(ROOT, 'tmp');
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const upload = multer({ dest: TMP, limits: { fileSize: 200 * 1024 * 1024 } });

for (const d of [DATA, BACKUPS, TMP, path.join(PUBLIC,'assets','uma'), path.join(PUBLIC,'assets','songs')]) fs.mkdirSync(d,{recursive:true});

app.use(express.json());
app.use(express.urlencoded({extended:true}));
app.use(express.static(PUBLIC, { etag:false, maxAge:0, setHeaders:(res)=>res.setHeader('Cache-Control','no-store') }));

function normalizeTarget(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  return path.posix.normalize(path.posix.join(base, target)).replace(/^\.\//,'');
}
function arr(v){ return v === undefined ? [] : (Array.isArray(v) ? v : [v]); }
function getAttr(obj,name){ return obj && obj['@_'+name]; }

async function parseAndPublish(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath, { cellFormula:true, cellText:true });
  const ws = wb.Sheets['楽曲データ'];
  if (!ws) throw new Error('「楽曲データ」シートがありません。');
  if (!wb.Sheets['ウマ娘']) throw new Error('「ウマ娘」シートがありません。');
  if (!wb.Sheets['楽曲']) throw new Error('「楽曲」シートがありません。');

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  const songs=[];
  for(let c=2;c<=range.e.c;c++){
    const v=ws[XLSX.utils.encode_cell({r:2,c})]?.v;
    if(typeof v==='string' && v.trim()) songs.push({c,name:v.trim()});
  }
  const umas=[]; const umaToSongs={}; const songToUmas={};
  for(const s of songs) songToUmas[s.name]=[];
  for(let r=3;r<=range.e.r;r++){
    const name=ws[XLSX.utils.encode_cell({r,c:1})]?.v;
    if(typeof name!=='string' || !name.trim()) continue;
    const n=name.trim(); const hits=[];
    for(const s of songs){
      const val=ws[XLSX.utils.encode_cell({r,c:s.c})]?.v;
      if(val!==undefined && val!==null && val!=='') { hits.push(s.name); songToUmas[s.name].push(n); }
    }
    umas.push(n); umaToSongs[n]=hits;
  }

  const zip = new AdmZip(xlsxPath);
  const xml = new XMLParser({ ignoreAttributes:false, attributeNamePrefix:'@_', removeNSPrefix:true, parseTagValue:false });
  const readXml = (entry)=>{
    const e=zip.getEntry(entry); if(!e) throw new Error(`Excel内部ファイルがありません: ${entry}`);
    return xml.parse(e.getData().toString('utf8'));
  };
  const workbook=readXml('xl/workbook.xml').workbook;
  const wbRels=readXml('xl/_rels/workbook.xml.rels').Relationships;
  const relMap={}; for(const r of arr(wbRels.Relationship)) relMap[getAttr(r,'Id')]=getAttr(r,'Target');
  const sheetMap={};
  for(const s of arr(workbook.sheets.sheet)){
    const rid=getAttr(s,'r:id') || getAttr(s,'id');
    let t=relMap[rid];
    if(!t) continue;
    sheetMap[getAttr(s,'name')]=normalizeTarget('xl',t);
  }

  const rvDoc=readXml('xl/richData/rdrichvalue.xml').rvData;
  const rvRelDoc=readXml('xl/richData/richValueRel.xml').richValueRels;
  const rvRelsDoc=readXml('xl/richData/_rels/richValueRel.xml.rels').Relationships;
  const relOrder=arr(rvRelDoc.rel).map(r=>getAttr(r,'id')||getAttr(r,'r:id'));
  const richTargets={}; for(const r of arr(rvRelsDoc.Relationship)) richTargets[getAttr(r,'Id')]=getAttr(r,'Target');
  const rvMedia=[];
  for(const rv of arr(rvDoc.rv)){
    const values=arr(rv.v).map(v=> typeof v==='object' ? (v['#text']??'') : v);
    const idx=parseInt(values[0],10);
    const rid=relOrder[idx];
    rvMedia.push(normalizeTarget('xl/richData', richTargets[rid] || ''));
  }

  function cellList(sheetName){
    const sh=readXml(sheetMap[sheetName]).worksheet;
    const rows=arr(sh.sheetData?.row); const cells=[];
    for(const row of rows) for(const c of arr(row.c)) cells.push(c);
    return cells;
  }
  function imageRowMap(sheetName){
    const m={};
    for(const c of cellList(sheetName)){
      const ref=getAttr(c,'r')||''; const vm=getAttr(c,'vm');
      if(/^C\d+$/.test(ref) && vm){ const row=parseInt(ref.slice(1),10); m[row]=rvMedia[parseInt(vm,10)-1]; }
    }
    return m;
  }
  const umaRows=imageRowMap('ウマ娘');
  const songRows=imageRowMap('楽曲');

  const stage=path.join(TMP,'publish_'+Date.now());
  await fsp.mkdir(path.join(stage,'uma'),{recursive:true});
  await fsp.mkdir(path.join(stage,'songs'),{recursive:true});
  const copyMedia=(media,dest)=>{
    if(!media) return null;
    const e=zip.getEntry(media); if(!e) return null;
    fs.writeFileSync(dest,e.getData()); return true;
  };
  const umaImages={};
  for(let i=0;i<umas.length;i++){
    const media=umaRows[i+3]; if(!media) continue;
    const ext=path.extname(media)||'.png'; const fn=`uma_${String(i+1).padStart(3,'0')}${ext}`;
    if(copyMedia(media,path.join(stage,'uma',fn))) umaImages[umas[i]]=`assets/uma/${fn}`;
  }
  const songImages={};
  for(let i=0;i<songs.length;i++){
    const media=songRows[i+3]; if(!media) continue;
    const ext=path.extname(media)||'.png'; const fn=`song_${String(i+1).padStart(3,'0')}${ext}`;
    if(copyMedia(media,path.join(stage,'songs',fn))) songImages[songs[i].name]=`assets/songs/${fn}`;
  }
  if(Object.keys(umaImages).length !== umas.length) throw new Error(`ウマ娘画像の対応に失敗しました (${Object.keys(umaImages).length}/${umas.length})`);
  if(Object.keys(songImages).length !== songs.length) throw new Error(`楽曲画像の対応に失敗しました (${Object.keys(songImages).length}/${songs.length})`);

  const payload={
    version:Date.now(),
    updatedAt:new Date().toISOString(),
    umas:umas.map(n=>({name:n,image:umaImages[n],songs:umaToSongs[n]||[]})),
    songs:songs.map(s=>({name:s.name,image:songImages[s.name],umas:songToUmas[s.name]||[]}))
  };
  const newJson=JSON.stringify(payload);
  JSON.parse(newJson);

  const current=path.join(DATA,'current.xlsx');
  if(fs.existsSync(current)){
    const stamp=new Date().toISOString().replace(/[:.]/g,'-');
    await fsp.copyFile(current,path.join(BACKUPS,`backup_${stamp}.xlsx`));
  }
  await fsp.rm(path.join(PUBLIC,'assets','uma'),{recursive:true,force:true});
  await fsp.rm(path.join(PUBLIC,'assets','songs'),{recursive:true,force:true});
  await fsp.mkdir(path.join(PUBLIC,'assets'),{recursive:true});
  await fsp.rename(path.join(stage,'uma'),path.join(PUBLIC,'assets','uma'));
  await fsp.rename(path.join(stage,'songs'),path.join(PUBLIC,'assets','songs'));
  await fsp.writeFile(path.join(PUBLIC,'data.json'),newJson,'utf8');
  await fsp.copyFile(xlsxPath,current);
  await fsp.rm(stage,{recursive:true,force:true});
  return {umas:umas.length,songs:songs.length,umaImages:Object.keys(umaImages).length,songImages:Object.keys(songImages).length};
}

app.post('/api/admin/update', upload.single('excel'), async (req,res)=>{
  const uploaded=req.file?.path;
  try{
    if(!ADMIN_PASSWORD) return res.status(503).json({ok:false,error:'ADMIN_PASSWORDがサーバーに設定されていません。'});
    if(req.body.password !== ADMIN_PASSWORD) return res.status(401).json({ok:false,error:'管理パスワードが違います。'});
    if(!req.file) return res.status(400).json({ok:false,error:'Excelファイルを選択してください。'});
    if(!/\.xlsx$/i.test(req.file.originalname)) return res.status(400).json({ok:false,error:'.xlsxファイルのみ対応しています。'});
    const result=await parseAndPublish(uploaded);
    res.json({ok:true,...result,message:'更新完了。同じURLのまま新しいデータに切り替わりました。'});
  } catch(e){ console.error(e); res.status(500).json({ok:false,error:e.message||String(e)}); }
  finally { if(uploaded) fsp.rm(uploaded,{force:true}).catch(()=>{}); }
});

app.get('/api/status', async (req,res)=>{
  try { const d=JSON.parse(await fsp.readFile(path.join(PUBLIC,'data.json'),'utf8')); res.json({ok:true,updatedAt:d.updatedAt||null,umas:d.umas.length,songs:d.songs.length}); }
  catch { res.json({ok:false}); }
});

app.get('/admin', (req,res)=>res.sendFile(path.join(PUBLIC,'admin.html')));
app.get('*',(req,res)=>res.sendFile(path.join(PUBLIC,'index.html')));
app.listen(PORT,()=>console.log(`Uma Song site: http://localhost:${PORT}`));
