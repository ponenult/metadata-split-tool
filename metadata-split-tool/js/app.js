// 全局变量
let raw = [];
let sep = ":";
let mode = "ltr";
let maxC = 0;
let cols = [];
let fname = "";
// 处理文件拖拽上传
function onDrop(e){ e.preventDefault(); const f=e.dataTransfer.files[0]; if(f) handleFile(f); }
// 读取用户上传的 Excel/CSV 文件，并初始化页面数据
function handleFile(file){
  if(!file) return;
  fname=file.name;
  const r=new FileReader();
  r.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1});
      raw=rows.map(r=>r[0]!==undefined?String(r[0]):'').filter(r=>r!=='');
      sep = detectSeparator(raw[0]);
      document.getElementById('fn').textContent=fname;
      document.getElementById('rc').textContent=raw.length+' 行';
      document.getElementById('config-section').style.display='block';
      //setSep(':',document.querySelector('.sep-btn'));
      const btn = [...document.querySelectorAll(".sep-btn")]
    .find(b => b.textContent.includes(sep));
    setSep(sep, btn);

    }catch(err){ showStatus('文件读取失败：'+err.message,'error'); }
  };
  r.readAsArrayBuffer(file);
}
// 设置分隔符，并重新解析数据
function setSep(s,btn){
  sep=s;
  document.querySelectorAll('.sep-btn').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  document.getElementById('cs').value=s;
  proc();
}
// 设置自定义分隔符
function setSepCustom(s){ if(!s) return; sep=s; document.querySelectorAll('.sep-btn').forEach(b=>b.classList.remove('active')); proc(); }
// 切换分列模式（从左填充 / 首尾对齐）
function setMode(m){
  mode=m;
  document.getElementById('mode-ltr').classList.toggle('active', m==='ltr');
  document.getElementById('mode-align').classList.toggle('active', m==='align');
  proc();
}
// 自动识别分隔符
function detectSeparator(text) {
    const separators = [":", "-", "/", "_"];

    for (const separator of separators) {
        if (text.includes(separator)) {
            return separator;
        }
    }

    return ":";
}
// 根据分隔符拆分一行数据
function splitRow(s){
  const parts=s.split(sep);
  const out=Array(maxC).fill('');
  if(mode==='ltr'){
    parts.slice(0,maxC).forEach((p,i)=>out[i]=p);
  } else {
    out[0]=parts[0]||'';
    out[maxC-1]=parts[parts.length-1]||'';
    for(let i=1;i<parts.length-1;i++) out[i]=parts[i];
  }
  return out;
}
// 重新计算数据并刷新预览
function proc(){
  if(!raw.length) return;
  maxC=Math.max(...raw.map(r=>r.split(sep).length));
  ensureCols(); renderCols();
  renderPrev(raw.slice(0,5).map(splitRow));
  // 更新统计信息
  updateQualityPanel();
}
// 根据最大列数生成默认列名
function ensureCols(){
  const def=['大类','分类维度','细分类别','指标名称','列5','列6','列7','列8'];
  const old=cols.slice(); cols=[];
  for(let i=0;i<maxC;i++) cols.push(old[i]||def[i]||('列'+(i+1)));
}
// 渲染列名输入框
function renderCols(){
  const c=document.getElementById('cn'); c.innerHTML='';
  cols.forEach((name,i)=>{
    const w=document.createElement('div'); w.className='col-item';
    const sp=document.createElement('span'); sp.textContent='列'+(i+1)+':';
    const inp=document.createElement('input'); inp.type='text'; inp.value=name; inp.style.width='90px';
    inp.oninput=function(){ cols[i]=this.value; renderPrev(raw.slice(0,5).map(splitRow)); };
    w.appendChild(sp); w.appendChild(inp); c.appendChild(w);
  });
}
// 渲染前 5 行预览数据
function renderPrev(rows){
  const t=document.getElementById('pt'); t.innerHTML='';
  const tr=t.createTHead().insertRow();
  cols.forEach(n=>{ const th=document.createElement('th'); th.textContent=n; tr.appendChild(th); });
  const tb=t.createTBody();
  rows.forEach(row=>{ const tr=tb.insertRow(); row.forEach(cell=>{ const td=tr.insertCell(); td.textContent=cell; }); });
}
// 导出分列后的 Excel 文件
function exportFile(){
  const wsData=[cols,...raw.map(splitRow)];
  const ws=XLSX.utils.aoa_to_sheet(wsData);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'分列结果');
  XLSX.writeFile(wb, fname.replace(/\.[^.]+$/,'')+'_分列结果.xlsx');
  showStatus('✅ 导出成功！','success');
}
// 显示操作状态提示
function showStatus(msg,type){
  const el=document.getElementById('sm');
  el.textContent=msg; el.className='status '+type; el.style.display='block';
  setTimeout(()=>el.style.display='none',3000);
}
//显示统计信息
function updateQualityPanel(){
 const qualityPanel = document.getElementById("qualityPanel");
 const duplicateCount = countDuplicate();
    qualityPanel.innerHTML = `
      最大层级：${maxC}
      重复数据：${duplicateCount}
      <p>空白行：--</p>
      <p>异常格式：--</p>
      
      
    `;
}
//统计重复数据
function countDuplicate() {


    const map = {};

    raw.forEach(item => {
        map[item] = (map[item] || 0) + 1;
    });

    let count = 0;

    Object.values(map).forEach(num => {
        if(num > 1){
            count++;
        }
    });

    return count;
}
