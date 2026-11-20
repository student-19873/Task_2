/* 산업연관분석 플랫폼
 * 입력 형식: Z 중간거래행렬(행=공급산업 i, 열=구매산업 j), 총산출액 x_j
 */
let workbook=null, sectorNames=[], Z=[], x=[], L=[], indicators=[], charts={};
const $=id=>document.getElementById(id);
const fmt=(v,d=3)=>Number.isFinite(v)?v.toLocaleString('ko-KR',{maximumFractionDigits:d}):'-';
const fmtMoney=(v)=>Number.isFinite(v)?v.toLocaleString('ko-KR',{maximumFractionDigits:2}):'-';
function status(msg){$('statusText').textContent=msg}
function setStatus(msg){status(msg)}
function colLetter(n){let s='';while(n>0){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s}
function parseNum(v){if(typeof v==='number')return v;if(v===null||v===undefined||v==='')return 0;const n=Number(String(v).replace(/,/g,''));return Number.isFinite(n)?n:NaN}
$('fileInput').addEventListener('change',e=>{
  const f=e.target.files[0];if(!f)return;
  $('fileName').textContent=f.name;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      workbook=XLSX.read(ev.target.result,{type:'array'});
      $('sheetSelect').innerHTML='';
      workbook.SheetNames.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=n;$('sheetSelect').appendChild(o)});
      $('sheetSelect').disabled=false;$('loadDataBtn').disabled=false;
      setStatus(`파일을 읽었습니다. ${workbook.SheetNames.length}개 시트 중 분석할 시트를 선택하세요.`);
    }catch(err){setStatus('엑셀 파일을 읽지 못했습니다: '+err.message)}
  };reader.readAsArrayBuffer(f);
});
$('loadDataBtn').addEventListener('click',()=>{try{loadWorkbookSheet()}catch(e){setStatus('분석 오류: '+e.message);console.error(e)}});
function loadWorkbookSheet(){
  const ws=workbook.Sheets[$('sheetSelect').value];if(!ws)throw Error('시트를 선택하세요.');
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  const h=Number($('headerRow').value)-1, r=Number($('firstRow').value)-1, c=Number($('firstCol').value)-1, n=Number($('sectorCount').value), outCol=Number($('outputCol').value)-1;
  if(!Number.isInteger(n)||n<1||r<0||c<0)throw Error('행·열·산업 개수를 확인하세요.');
  if(rows.length<r+n)throw Error('지정한 행 범위가 시트 범위를 벗어났습니다.');
  sectorNames=Array.from({length:n},(_,j)=>String(rows[h]?.[c+j]??`산업 ${j+1}`));
  Z=Array.from({length:n},(_,i)=>Array.from({length:n},(_,j)=>{
    const v=parseNum(rows[r+i]?.[c+j]);if(!Number.isFinite(v))throw Error(`${r+i+1}행 ${c+j+1}열에 숫자가 아닌 값이 있습니다.`);return v;
  }));
  if(outCol>=0){x=Array.from({length:n},(_,i)=>parseNum(rows[r+i]?.[outCol]));if(x.some(v=>!Number.isFinite(v)||v<=0))throw Error('총산출액 열에 0 이하 또는 숫자가 아닌 값이 있습니다.')}
  else{x=Array.from({length:n},(_,j)=>Z.reduce((s,row)=>s+row[j],0));if(x.some(v=>v<=0))throw Error('열합으로 총산출액을 계산할 수 없습니다. 총산출액 열을 지정하세요.')}
  analyze();setStatus(`${$('sheetSelect').value} 시트 분석 완료 · ${n}개 산업 · 단위 ${$('unitLabel').value||'미지정'}`);
}
function identity(n){return Array.from({length:n},(_,i)=>Array.from({length:n},(_,j)=>i===j?1:0))}
function inverse(a){
  const n=a.length,m=a.map((row,i)=>[...row,...identity(n)[i]]);
  for(let col=0;col<n;col++){
    let pivot=col;for(let r=col+1;r<n;r++)if(Math.abs(m[r][col])>Math.abs(m[pivot][col]))pivot=r;
    if(Math.abs(m[pivot][col])<1e-12)throw Error('행렬이 역행렬을 갖지 않습니다. 산업 범위와 데이터 방향을 확인하세요.');
    [m[col],m[pivot]]=[m[pivot],m[col]];
    const div=m[col][col];for(let j=0;j<2*n;j++)m[col][j]/=div;
    for(let r=0;r<n;r++){if(r===col)continue;const factor=m[r][col];for(let j=0;j<2*n;j++)m[r][j]-=factor*m[col][j]}
  }return m.map(row=>row.slice(n));
}
function analyze(){
  const n=Z.length,A=Z.map(row=>row.map((v,j)=>v/x[j]));
  L=inverse(identity(n).map((row,i)=>row.map((v,j)=>v-A[i][j])));
  const rowSums=Array(n).fill(0),colSums=Array(n).fill(0);
  for(let i=0;i<n;i++)for(let j=0;j<n;j++){rowSums[i]+=L[i][j];colSums[j]+=L[i][j]}
  const avg=colSums.reduce((a,b)=>a+b,0)/n;
  indicators=sectorNames.map((name,i)=>({name,production:colSums[i],influence:colSums[i]/avg,sensitivity:rowSums[i]/avg,output:x[i]}));
  render();
}
function render(){
  const n=sectorNames.length,unit=$('unitLabel').value||'미지정';
  $('kpiCount').textContent=n;$('kpiProduction').textContent=fmt(indicators.reduce((s,d)=>s+d.production,0)/n);
  $('kpiInfluence').textContent=fmt(indicators.reduce((s,d)=>s+d.influence,0)/n);
  $('kpiSensitivity').textContent=fmt(indicators.reduce((s,d)=>s+d.sensitivity,0)/n);
  $('unitPill').textContent='단위: '+unit;
  $('resultsBody').innerHTML=indicators.map(d=>`<tr><td><strong>${esc(d.name)}</strong></td><td>${fmt(d.production)}</td><td>${fmt(d.influence)}</td><td>${fmt(d.sensitivity)}</td><td>${fmtMoney(d.output)} ${esc(unit)}</td></tr>`).join('');
  const ranked=[...indicators].sort((a,b)=>b.influence-a.influence);
  $('rankingBody').innerHTML=ranked.map((d,i)=>`<tr><td>${i+1}</td><td><strong>${esc(d.name)}</strong></td><td>${fmt(d.influence)}</td><td>${fmt(d.sensitivity)}</td><td>${d.influence>=1?'평균보다 큰 전방위 파급': '평균보다 낮은 영향력'}</td></tr>`).join('');
  $('simSector').innerHTML=sectorNames.map((s,i)=>`<option value="${i}">${esc(s)}</option>`).join('');$('simSector').disabled=false;$('simulateBtn').disabled=false;
  drawChart('productionChart',{type:'bar',data:{labels:sectorNames,datasets:[{label:'생산유발계수',data:indicators.map(d=>d.production),backgroundColor:'#3977d5',borderRadius:5} ]},options:chartOptions()});
  drawChart('linkageChart',{type:'bar',data:{labels:sectorNames,datasets:[{label:'영향력계수',data:indicators.map(d=>d.influence),backgroundColor:'#3977d5'},{label:'감응도계수',data:indicators.map(d=>d.sensitivity),backgroundColor:'#18a6a0'}]},options:chartOptions()});
  $('simResults').innerHTML='<div class="empty">산업과 최종수요 증가액을 선택한 뒤 계산을 실행하세요.</div>';
}
function chartOptions(){return{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{boxWidth:10,font:{family:'Noto Sans KR'}}}},scales:{x:{ticks:{font:{size:10}}},y:{beginAtZero:true}}}}
function drawChart(id,config){if(charts[id])charts[id].destroy();charts[id]=new Chart($(id),config)}
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
$('simulateBtn').addEventListener('click',()=>{
  if(!L.length)return;
  const j=Number($('simSector').value),shock=Number($('shockAmount').value);
  if(!Number.isFinite(shock)||shock<0){setStatus('최종수요 증가액은 0 이상의 숫자로 입력하세요.');return}
  const delta=Array(sectorNames.length).fill(0);delta[j]=shock;
  const effects=L.map(row=>row.reduce((s,v,k)=>s+v*delta[k],0));
  const total=effects.reduce((a,b)=>a+b,0);
  $('simResults').innerHTML=`<div class="sim-metric"><span>총 생산유발액</span><strong>${fmtMoney(total)} ${esc($('simUnit').value)}</strong></div><div class="sim-metric"><span>직접 최종수요 증가액</span><strong>${fmtMoney(shock)} ${esc($('simUnit').value)}</strong></div><div class="sim-metric"><span>간접 생산유발액</span><strong>${fmtMoney(total-shock)} ${esc($('simUnit').value)}</strong></div>`;
  drawChart('simulationChart',{type:'bar',data:{labels:sectorNames,datasets:[{label:'산업별 생산유발액',data:effects,backgroundColor:'#159b91',borderRadius:5}]},options:chartOptions()});
  setStatus(`${sectorNames[j]} 산업의 최종수요 ${fmtMoney(shock)} ${$('simUnit').value} 증가를 가정한 결과입니다.`);
});
$('sampleBtn').addEventListener('click',()=>{
  sectorNames=['농림수산품','광산품','제조업','전력·가스','서비스'];
  Z=[[20,5,30,2,8],[2,3,12,1,3],[15,10,180,25,45],[3,2,20,15,18],[12,8,55,20,95]];
  x=[100,80,500,120,300];$('unitLabel').value='억원';$('simUnit').value='억원';analyze();
  setStatus('내장 예시 데이터로 분석 중입니다. 실제 과제 데이터가 아닌 기능 시연용 가상 데이터입니다.');
});
$('downloadSampleBtn').addEventListener('click',()=>{
  const data=[['산업','농림수산품','광산품','제조업','전력·가스','서비스','총산출액'],
  ['농림수산품',20,5,30,2,8,100],['광산품',2,3,12,1,3,80],['제조업',15,10,180,25,45,500],['전력·가스',3,2,20,15,18,120],['서비스',12,8,55,20,95,300]];
  const ws=XLSX.utils.aoa_to_sheet(data),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'산업연관표');XLSX.writeFile(wb,'산업연관표_예시데이터.xlsx');
});
