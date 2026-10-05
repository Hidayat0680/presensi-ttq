let santri = JSON.parse(localStorage.getItem('FINAL_IDPPS') || '[]');
let absen = JSON.parse(localStorage.getItem('FINAL_ABSEN') || '[]');

function showPage(p){
  document.querySelectorAll('[id^="page-"]').forEach(e=>e.classList.add('hidden'));
  let page = document.getElementById('page-'+p);
  if(page) page.classList.remove('hidden');
  document.querySelectorAll('.menu a').forEach(e=>e.classList.remove('active'));
  let el = document.getElementById('m-'+p);
  if(el) el.classList.add('active');
  if(p==='absen'){
    setTimeout(()=>{
      let i=document.getElementById('kodeAbsen');
      if(i) i.focus();
    },200);
  }
  load();
}

function save(){
  localStorage.setItem('FINAL_IDPPS', JSON.stringify(santri));
  localStorage.setItem('FINAL_ABSEN', JSON.stringify(absen));
  load();
}

function load(){
  // Tabel Data
  let tabel = document.getElementById('tabel');
  if(tabel){
    tabel.innerHTML = santri.slice(0,300).map(s=>`<tr><td><span class="code">${s.idpps}</span></td><td>${s.nama}</td></tr>`).join('') || '<tr><td colspan=2 style="text-align:center;padding:20px;color:#aaa">Belum ada data - Import Excel dulu</td></tr>';
  }
  let info = document.getElementById('infoCount');
  if(info) info.innerText = santri.length+' santri';

  let sTotal = document.getElementById('s-total');
  if(sTotal) sTotal.innerText = santri.length;

  let today = new Date().toLocaleDateString('id-ID');
  let todayAbsen = absen.filter(a=>a.tgl===today);

  let sHadir = document.getElementById('s-hadir');
  if(sHadir) sHadir.innerText = todayAbsen.filter(a=>a.status==='Hadir').length;
  let sIzin = document.getElementById('s-izin');
  if(sIzin) sIzin.innerText = todayAbsen.filter(a=>a.status==='Izin').length;
  let sAlpha = document.getElementById('s-alpha');
  if(sAlpha) sAlpha.innerText = Math.max(0, santri.length - todayAbsen.length);

  let list = document.getElementById('listAbsen');
  if(list){
    list.innerHTML = todayAbsen.slice(-5).reverse().map(a=>`<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #eee"><span><span class="code" style="font-size:11px">${a.idpps}</span> ${a.nama}</span><span>${a.jam} <b style="color:green">${a.status}</b></span></div>`).join('') || 'Belum ada absen hari ini. Scan di menu Presensi!';
  }

  let lap = document.getElementById('tabelLap');
  if(lap){
    lap.innerHTML = absen.slice().reverse().map(a=>`<tr><td>${a.tgl}</td><td><b>${a.jam}</b><br><small style="color:#65a30d">${a.jamWIS||''}</small></td><td><span class="code">${a.idpps}</span></td><td>${a.nama}</td><td><span style="background:#dcfce7;padding:3px 8px;border-radius:10px;font-size:11px">${a.status}</span></td></tr>`).join('') || '<tr><td colspan=5 style="text-align:center">Belum ada</td></tr>';
  }
}

function importExcel(e){
  let file = e.target.files[0];
  if(!file) return;
  let reader = new FileReader();
  reader.onload = function(evt){
    let wb = XLSX.read(new Uint8Array(evt.target.result),{type:'array'});
    let ws = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(ws,{header:1});
    let start = json[0] && String(json[0][0]).toUpperCase().includes('IDPPS')? 1 : 0;
    let c = 0;
    for(let i=start;i<json.length;i++){
      let r = json[i];
      if(!r||!r[0]) continue;
      let id = String(r[0]).split('.')[0].trim();
      if(!id || santri.find(s=>String(s.idpps)===id)) continue;
      santri.push({idpps:id, nama: r[1]?String(r[1]).trim():'-'});
      c++;
    }
    save();
    alert('Import '+c+' santri berhasil! IDPPS jadi kode absen. ENTER langsung Hadir!');
  };
  reader.readAsArrayBuffer(file);
}

function downloadTemplate(){
  let ws = XLSX.utils.aoa_to_sheet([['IDPPS','Nama'],['39058','MOCH ABID'],['14370680','M. HIDAYATULLOH']]);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Data');
  XLSX.writeFile(wb,'Template-IDPPS.xlsx');
}

function hapusSemua(){
  if(confirm('Hapus semua data santri & absen?')){
    santri=[]; absen=[];
    localStorage.setItem('FINAL_IDPPS','[]');
    localStorage.setItem('FINAL_ABSEN','[]');
    save();
  }
}

function prosesAbsen(status){
  let el = document.getElementById('kodeAbsen');
  if(!el) return;
  let kode = String(el.value.trim().split('.')[0]);
  if(!kode) return;
  let s = santri.find(x=>String(x.idpps)===kode);
  let h = document.getElementById('hasil');

  if(!s){
    if(h){h.style.display='block';h.style.background='#fee2e2';h.innerHTML='❌ IDPPS '+kode+' tidak ditemukan!';}
    el.select(); return;
  }
  let today = new Date().toLocaleDateString('id-ID');
  let now = new Date();
  let jamWIB = now.toLocaleTimeString('id-ID',{hour12:false})+' WIB';
  let jamWIS = new Date(now.getTime()+43*60000).toLocaleTimeString('id-ID',{hour12:false})+' WIS';

  if(absen.find(a=>String(a.idpps)===kode && a.tgl===today)){
    let sudah = absen.find(a=>String(a.idpps)===kode && a.tgl===today);
    if(h){h.style.display='block';h.style.background='#fef3c7';h.innerHTML='⚠️ '+s.nama+' sudah absen jam '+sudah.jam;}
    el.value=''; el.focus(); return;
  }
  absen.push({idpps:s.idpps, nama:s.nama, status, tgl:today, jam:jamWIB, jamWIS});
  save();
  if(h){
    h.style.display='block';h.style.background='#dcfce7';
    h.innerHTML='✅ Berhasil!<br><b>'+s.nama+'</b> (IDPPS: '+s.idpps+')<br>Jam: '+jamWIB+' | '+jamWIS+'<br>Status: '+status;
  }
  el.value=''; el.focus();
}

function exportExcel(){
  let ws = XLSX.utils.aoa_to_sheet([['Tanggal','Jam WIB','Jam WIS','IDPPS','Nama','Status'],...absen.map(a=>[a.tgl,a.jam,a.jamWIS||'',a.idpps,a.nama,a.status])]);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Rekap');
  XLSX.writeFile(wb,'Rekap-Sidogiri-'+new Date().toISOString().slice(0,10)+'.xlsx');
}

// INI KUNCINYA: ENTER = HADIR
document.addEventListener('DOMContentLoaded', function(){
  let input = document.getElementById('kodeAbsen');
  if(input){
    input.addEventListener('keydown', function(e){
      if(e.key==='Enter'){
        e.preventDefault();
        prosesAbsen('Hadir');
      }
    });
  }
  load();

  // Jam WIB & WIS Realtime
  setInterval(function(){
    let now = new Date();
    let wib = now.toLocaleTimeString('id-ID',{hour12:false,timeZone:'Asia/Jakarta'});
    let wis = new Date(now.getTime()+43*60000).toLocaleTimeString('id-ID',{hour12:false,timeZone:'Asia/Jakarta'});
    let elWIB = document.getElementById('jamWIB'); if(elWIB) elWIB.innerText = wib+' WIB';
    let elWIS = document.getElementById('jamWIS'); if(elWIS) elWIS.innerText = wis+' WIS';
    let elLive = document.getElementById('jamLive'); if(elLive) elLive.innerText = wib+' / '+wis+' WIS';
    let elTgl = document.getElementById('tglNow'); if(elTgl) elTgl.innerText = now.toLocaleDateString('id-ID',{weekday:'long',day:'2-digit',month:'long',year:'numeric'});
  },1000);
});
