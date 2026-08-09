let DATA={umas:[],songs:[]}; let mode='uma'; let subset=null; let selected='';
const grid=document.getElementById('grid'), search=document.getElementById('search'), count=document.getElementById('count'), selection=document.getElementById('selection');
const tabs=[...document.querySelectorAll('.tab')];
function list(){ return subset ?? (mode==='uma'?DATA.umas:DATA.songs); }
function render(){
  const q=search.value.trim().toLowerCase(); const items=list().filter(x=>x.name.toLowerCase().includes(q));
  count.textContent=`${items.length}件`;
  grid.innerHTML=items.length?'':'<div class="empty">該当するデータがありません</div>';
  for(const item of items){ const b=document.createElement('button'); b.className='card'; b.innerHTML=`<img loading="lazy" src="${item.image}" alt=""><div class="name"></div>`; b.querySelector('.name').textContent=item.name; b.onclick=()=>choose(item); grid.appendChild(b); }
}
function choose(item){
  selected=item.name;
  if(mode==='uma'){
    const set=new Set(item.songs); subset=DATA.songs.filter(s=>set.has(s.name)); mode='song';
    selection.innerHTML=`選択したウマ娘：<strong></strong> ／ 歌唱楽曲 ${subset.length}曲`; selection.querySelector('strong').textContent=item.name;
  }else{
    const set=new Set(item.umas); subset=DATA.umas.filter(u=>set.has(u.name)); mode='uma';
    selection.innerHTML=`選択した楽曲：<strong></strong> ／ 歌唱ウマ娘 ${subset.length}名`; selection.querySelector('strong').textContent=item.name;
  }
  selection.classList.remove('hidden'); search.value=''; syncTabs(); render(); window.scrollTo({top:0,behavior:'smooth'});
}
function syncTabs(){ tabs.forEach(t=>t.classList.toggle('active',t.dataset.mode===mode)); search.placeholder=mode==='uma'?'ウマ娘名を検索':'楽曲名を検索'; }
function reset(newMode=mode){ mode=newMode; subset=null; selected=''; selection.classList.add('hidden'); search.value=''; syncTabs(); render(); }
tabs.forEach(t=>t.onclick=()=>reset(t.dataset.mode)); search.addEventListener('input',render); document.getElementById('homeBtn').onclick=()=>reset(mode);
fetch('data.json?ts='+Date.now(),{cache:'no-store'}).then(r=>r.json()).then(d=>{DATA=d;render()}).catch(()=>grid.innerHTML='<div class="empty">データの読み込みに失敗しました</div>');
