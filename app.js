let defaultUsers=[{username:'admin',password:'admin123',nama:'Admin Utama',role:'admin'}];
let users=JSON.parse(localStorage.getItem('ttq_users')||'null')||defaultUsers;
let currentUser=JSON.parse(localStorage.getItem('ttq_currentUser')||'null');
let santri=JSON.parse(localStorage.getItem('ttq_santri')||'[]');
let presensi=JSON.parse(localStorage.getItem('ttq_presensi')||'[]');
function saveUsers(){localStorage.setItem('ttq_users',JSON.stringify(users));}
function save(){localStorage.setItem('ttq_santri',JSON.stringify(santri));localStorage.setItem('ttq_presensi',JSON.stringify(presensi));}
function checkLogin(){
  if(currentUser){
    document.getElementById('loginScreen').style.display='none';
    document.getElementById('app').style.display='block';
    document.getElementById('infoUser').innerText=currentUser.nama;
    document.getElementById('infoRole').innerText=currentUser.role;
    if(currentUser.role!=='admin'){document.getElementById('tabPenggunaBtn').style.display='none';}
    else{document.getElementById('tabPenggunaBtn').style.display='block';}
  } else {
    document.getElementById('loginScreen').style.display='flex';
    document.getElementById('app').style.display='none';
  }
}
function doLogin(){
  let u=document.getElementById('loginUser').value.trim();
  let p=document.getElementById('loginPass').value.trim();
  let found=users.find(x=>x.username===u && x.password===p);
  if(!found){document.getElementById('loginError').innerText='Username / Password salah!';return;}
  currentUser=found;
  localStorage.setItem('ttq_currentUser',JSON.stringify(currentUser));
  document.getElementById('loginError').innerText='';
  checkLogin();renderMaster();renderHariIni();renderLaporan();renderPengguna();
}
function doLogout(){
  if(!confirm('Logout?'))return;
  currentUser=null;localStorage.removeItem('ttq_currentUser');checkLogin();
}
let lu=document.getElementById('loginUser');
let lp=document.getElementById('loginPass');
if(lu) lu.addEventListener('keydown',e=>{if(e.key==='Enter')doLogin();});
if(lp) lp.addEventListener('keydown',e=>{if(e.key==='Enter')doLogin();});
function openTab(n){
  if(n==='pengguna' && currentUser && currentUser.role!=='admin'){alert('Hanya Admin!');return;}
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
  if(n==='presensi')document.querySelectorAll('.tab')[0].classList.add('active');
  if(n==='master')document.querySelectorAll('.tab')[1].classList.add('active');
  if(n==='laporan')document.querySelectorAll('.tab')[2].classList.add('active');
  if(n==='pengguna')document.getElementById('tabPenggunaBtn').classList.add('active');
  document.getElementById('tab-'+n).classList.add('active');
  if(n==='laporan')renderLaporan();if(n==='master')renderMaster();if(n==='presensi')renderHariIni();if(n==='pengguna')renderPengguna();
}
document.getElementById('tglHariIni').innerText=new Date().toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
document.getElementById('filterTanggal').valueAsDate=new Date();
document.getElementById('fileExcel').addEventListener('change',e=>{
  let file=e.target.files[0];if(!file)return;
  let reader=new FileReader();
  reader.onload=function(ev){
    let wb=XLSX.read(ev.target.result,{type:'binary'});
    let rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    let baru=0;
    rows.forEach(r=>{
      let id=(r.IDPPS||r.Idpps||r.ID||'').toString().trim();
      let nama=(r.Nama||r.NAMA||'').toString().trim();
      if(id&&nama&&!santri.find(s=>s.id==id)){santri.push({id,nama});baru++;}
    });
    save();renderMaster();alert('Import '+baru+' santri. Total '+santri.length);
  };
  reader.readAsBinaryString(file);
});
document.getElementById('inputIDPPS').addEventListener('keydown',e=>{
  if(e.key==='Enter'){
    let id=e.target.value.trim();if(!id)return;
    let s=santri.find(x=>x.id==id);
    let res=document.getElementById('scanResult');
    if(!s){res.innerHTML='IDPPS '+id+' tidak ditemukan';e.target.value='';return;}
    let today=new Date().toISOString().slice(0,10);
    let sudah=presensi.find(p=>p.id==id&&p.tanggal==today);
    if(sudah){res.innerHTML=s.nama+' sudah hadir jam '+sudah.jam;}
    else{
      let jam=new Date().toLocaleTimeString('id-ID');
      presensi.push({id,nama:s.nama,tanggal:today,jam,petugas:currentUser.username});save();
      res.innerHTML=s.nama+' - HADIR jam '+jam;
    }
    e.target.value='';renderHariIni();renderLaporan();
  }
});
function renderHariIni(){
  let today=new Date().toISOString().slice(0,10);
  let data=presensi.filter(p=>p.tanggal==today).reverse();
  document.getElementById('jumlahHadir').innerText=data.length+' Hadir';
  document.getElementById('tabelHadirHariIni').innerHTML=data.map(p=>'<tr><td>'+p.jam+'</td><td>'+p.id+'</td><td>'+p.nama+'</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatu(\''+p.id+'\',\''+p.tanggal+'\')">Hapus</button></td></tr>').join('')||'<tr><td colspan=4>Belum ada</td></tr>';
}
function renderMaster(){
  let cari=document.getElementById('cariSantri'); let q=cari?cari.value.toLowerCase():'';
  let filtered=santri.filter(s=>s.id.toLowerCase().includes(q)||s.nama.toLowerCase().includes(q));
  document.getElementById('totalSantri').innerText=santri.length;
  document.getElementById('tabelMaster').innerHTML=filtered.map(s=>'<tr><td>'+s.id+'</td><td>'+s.nama+'</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatuSantri(\''+s.id+'\')">Hapus</button></td></tr>').join('');
}
function renderLaporan(){
  let t=document.getElementById('filterTanggal').value;
  let data=t?presensi.filter(p=>p.tanggal==t):presensi;
  document.getElementById('tabelLaporan').innerHTML=data.slice().reverse().map(p=>'<tr><td>'+p.tanggal+'</td><td>'+p.id+'</td><td>'+p.nama+'</td><td>'+p.jam+'</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatu(\''+p.id+'\',\''+p.tanggal+'\')">Hapus</button></td></tr>').join('')||'<tr><td colspan=5>Tidak ada data</td></tr>';
}
document.getElementById('filterTanggal').addEventListener('change',renderLaporan);
function renderPengguna(){
  document.getElementById('tabelPengguna').innerHTML=users.map(u=>'<tr><td>'+u.username+'</td><td>'+u.nama+'</td><td>'+u.role+'</td><td>'+(u.username==='admin'?'Default':'<button class="btn btn-danger btn-sm" onclick="hapusUser(\''+u.username+'\')">Hapus</button>')+'</td></tr>').join('');
}
function tambahUser(){
  let u=document.getElementById('newUser').value.trim();
  let n=document.getElementById('newNama').value.trim();
  let p=document.getElementById('newPass').value.trim();
  let r=document.getElementById('newRole').value;
  if(!u||!n||!p) return alert('Lengkapi semua!');
  if(users.find(x=>x.username===u)) return alert('Username sudah ada!');
  users.push({username:u,nama:n,password:p,role:r});saveUsers();renderPengguna();
  document.getElementById('newUser').value='';document.getElementById('newNama').value='';document.getElementById('newPass').value='';
  alert('Pengguna '+n+' ditambahkan!');
}
function hapusUser(username){
  if(username==='admin') return alert('Tidak bisa hapus admin default!');
  if(!confirm('Hapus '+username+'?')) return;
  users=users.filter(u=>u.username!==username);saveUsers();renderPengguna();
}
function hapusSatu(id,tanggal){if(!confirm('Hapus '+id+'?'))return;presensi=presensi.filter(p=>!(p.id==id&&p.tanggal==tanggal));save();renderHariIni();renderLaporan();}
function hapusSatuSantri(id){if(currentUser.role!=='admin') return alert('Hanya Admin!');if(!confirm('Hapus santri '+id+'?'))return;santri=santri.filter(s=>s.id!=id);presensi=presensi.filter(p=>p.id!=id);save();renderMaster();renderHariIni();renderLaporan();}
function hapusHariIni(){let today=new Date().toISOString().slice(0,10);let jml=presensi.filter(p=>p.tanggal==today).length;if(jml==0)return alert('Tidak ada data hari ini');if(!confirm('Hapus '+jml+' presensi hari ini?'))return;presensi=presensi.filter(p=>p.tanggal!=today);save();renderHariIni();renderLaporan();}
function hapusTanggalTerpilih(){let t=document.getElementById('filterTanggal').value;if(!t)return alert('Pilih tanggal dulu');let jml=presensi.filter(p=>p.tanggal==t).length;if(jml==0)return alert('Tidak ada data');if(!confirm('Hapus '+jml+' data tanggal '+t+'?'))return;presensi=presensi.filter(p=>p.tanggal!=t);save();renderHariIni();renderLaporan();}
function hapusSemuaPresensi(){if(currentUser.role!=='admin') return alert('Hanya Admin!');if(presensi.length==0)return alert('Sudah kosong');if(!confirm('YAKIN hapus SEMUA?'))return;if(!confirm('Konfirmasi kedua?'))return;presensi=[];save();renderHariIni();renderLaporan();}
function hapusSemuaSantri(){if(currentUser.role!=='admin') return alert('Hanya Admin!');if(santri.length==0)return alert('Sudah kosong');if(!confirm('YAKIN hapus SEMUA?'))return;if(!confirm('Konfirmasi kedua?'))return;santri=[];presensi=[];save();renderMaster();renderHariIni();renderLaporan();}
function exportExcel(){let t=document.getElementById('filterTanggal').value;let data=t?presensi.filter(p=>p.tanggal==t):presensi;if(data.length==0)return alert('Tidak ada data!');try{let ws=XLSX.utils.json_to_sheet(data);let wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Laporan');XLSX.writeFile(wb,'Laporan_TTQ_'+(t||'semua')+'.xlsx');}catch(e){alert('Export gagal: '+e.message);}}
checkLogin();renderMaster();renderHariIni();renderLaporan();renderPengguna();
