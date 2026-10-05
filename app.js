let defaultUsers=[{username:'admin',password:'admin123',nama:'Admin Utama',role:'admin'}];
let users=JSON.parse(localStorage.getItem('ttq_users')||'null');
if(!users ||!Array.isArray(users) || users.length===0){users=defaultUsers;localStorage.setItem('ttq_users',JSON.stringify(users));}
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
    document.getElementById('tabPenggunaBtn').style.display=currentUser.role==='admin'?'block':'none';
  }else{
    document.getElementById('loginScreen').style.display='flex';
    document.getElementById('app').style.display='none';
  }
}
function doLogin(){
  let u=document.getElementById('loginUser').value.trim();
  let p=document.getElementById('loginPass').value.trim();
  if(!u||!p){document.getElementById('loginError').innerText='Isi username & password!';return;}
  if(u==='admin' && p==='admin123'){
    let found=users.find(x=>x.username==='admin');
    if(!found){users.unshift({username:'admin',password:'admin123',nama:'Admin Utama',role:'admin'});saveUsers();}
    else{found.password='admin123';found.role='admin';saveUsers();}
    currentUser=users.find(x=>x.username==='admin');
    localStorage.setItem('ttq_currentUser',JSON.stringify(currentUser));
    checkLogin();renderMaster();renderHariIni();renderLaporan();renderPengguna();return;
  }
  let found=users.find(x=>x.username===u && x.password===p);
  if(!found){document.getElementById('loginError').innerText='GAGAL! Coba klik RESET di bawah!';return;}
  currentUser=found;localStorage.setItem('ttq_currentUser',JSON.stringify(currentUser));checkLogin();renderMaster();renderHariIni();renderLaporan();renderPengguna();
}
function resetLogin(){
  if(!confirm('Reset semua akun login ke default admin/admin123? Data santri aman.'))return;
  localStorage.removeItem('ttq_users');localStorage.removeItem('ttq_currentUser');
  users=[{username:'admin',password:'admin123',nama:'Admin Utama',role:'admin'}];
  saveUsers();currentUser=null;checkLogin();
  document.getElementById('loginUser').value='admin';document.getElementById('loginPass').value='admin123';
  document.getElementById('loginError').innerText='SUDAH DI RESET! Sekarang klik LOGIN atau tekan ENTER';
}
function doLogout(){if(!confirm('Logout?'))return;currentUser=null;localStorage.removeItem('ttq_currentUser');checkLogin();}
let lu=document.getElementById('loginUser');let lp=document.getElementById('loginPass');
if(lu)lu.addEventListener('keydown',e=>{if(e.key==='Enter')doLogin();});
if(lp)lp.addEventListener('keydown',e=>{if(e.key==='Enter')doLogin();});
function openTab(n){
  if(n==='pengguna'&&currentUser&&currentUser.role!=='admin'){alert('Hanya admin');return;}
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
  let tabs=document.querySelectorAll('.tab');
  if(n==='presensi')tabs[0].classList.add('active');if(n==='master')tabs[1].classList.add('active');if(n==='laporan')tabs[2].classList.add('active');if(n==='pengguna')document.getElementById('tabPenggunaBtn').classList.add('active');
  document.getElementById('tab-'+n).classList.add('active');
  if(n==='laporan')renderLaporan();if(n==='master')renderMaster();if(n==='presensi')renderHariIni();if(n==='pengguna')renderPengguna();
}
let tglEl=document.getElementById('tglHariIni');if(tglEl)tglEl.innerText=new Date().toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
let filt=document.getElementById('filterTanggal');if(filt)filt.valueAsDate=new Date();
let fe=document.getElementById('fileExcel');if(fe)fe.addEventListener('change',e=>{
  let file=e.target.files[0];if(!file)return;let reader=new FileReader();
  reader.onload=function(ev){let wb=XLSX.read(ev.target.result,{type:'binary'});let rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);let baru=0;rows.forEach(r=>{let id=(r.IDPPS||r.Idpps||r.ID||'').toString().trim();let nama=(r.Nama||r.NAMA||'').toString().trim();if(id&&nama&&!santri.find(s=>s.id==id)){santri.push({id,nama});baru++;}});save();renderMaster();alert('Import '+baru);};
  reader.readAsBinaryString(file);
});
let inp=document.getElementById('inputIDPPS');if(inp)inp.addEventListener('keydown',e=>{if(e.key==='Enter'){let id=e.target.value.trim();if(!id)return;let s=santri.find(x=>x.id==id);let res=document.getElementById('scanResult');if(!s){res.innerText='IDPPS '+id+' tidak ada';e.target.value='';return;}let today=new Date().toISOString().slice(0,10);let sudah=presensi.find(p=>p.id==id&&p.tanggal==today);if(sudah){res.innerText=s.nama+' sudah hadir '+sudah.jam;}else{let jam=new Date().toLocaleTimeString('id-ID');presensi.push({id,nama:s.nama,tanggal:today,jam,petugas:currentUser.username});save();res.innerText=s.nama+' HADIR '+jam;}e.target.value='';renderHariIni();renderLaporan();}});
function renderHariIni(){let today=new Date().toISOString().slice(0,10);let data=presensi.filter(p=>p.tanggal==today).reverse();document.getElementById('jumlahHadir').innerText=data.length+' Hadir';document.getElementById('tabelHadirHariIni').innerHTML=data.map(p=>`<tr><td>${p.jam}</td><td>${p.id}</td><td>${p.nama}</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatu('${p.id}','${p.tanggal}')">Hapus</button></td></tr>`).join('')||'<tr><td colspan=4>Belum ada</td></tr>';}
function renderMaster(){let q=(document.getElementById('cariSantri')?.value||'').toLowerCase();let f=santri.filter(s=>s.id.toLowerCase().includes(q)||s.nama.toLowerCase().includes(q));document.getElementById('totalSantri').innerText=santri.length;document.getElementById('tabelMaster').innerHTML=f.map(s=>`<tr><td>${s.id}</td><td>${s.nama}</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatuSantri('${s.id}')">Hapus</button></td></tr>`).join('');}
function renderLaporan(){let t=document.getElementById('filterTanggal').value;let data=t?presensi.filter(p=>p.tanggal==t):presensi;document.getElementById('tabelLaporan').innerHTML=data.slice().reverse().map(p=>`<tr><td>${p.tanggal}</td><td>${p.id}</td><td>${p.nama}</td><td>${p.jam}</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatu('${p.id}','${p.tanggal}')">Hapus</button></td></tr>`).join('')||'<tr><td colspan=5>Tidak ada</td></tr>';}
document.getElementById('filterTanggal')?.addEventListener('change',renderLaporan);
function renderPengguna(){document.getElementById('tabelPengguna').innerHTML=users.map(u=>`<tr><td>${u.username}</td><td>${u.nama}</td><td>${u.role}</td><td>${u.username==='admin'?'Default':'<button class="btn btn-danger btn-sm" onclick="hapusUser(\''+u.username+'\')">Hapus</button>'}</td></tr>`).join('');}
function tambahUser(){let u=document.getElementById('newUser').value.trim();let n=document.getElementById('newNama').value.trim();let p=document.getElementById('newPass').value.trim();let r=document.getElementById('newRole').value;if(!u||!n||!p)return alert('Lengkapi!');if(users.find(x=>x.username===u))return alert('Sudah ada!');users.push({username:u,nama:n,password:p,role:r});saveUsers();renderPengguna();alert('Ditambah!');}
function hapusUser(un){if(un==='admin')return alert('Tidak bisa');if(!confirm('Hapus '+un+'?'))return;users=users.filter(u=>u.username!==un);saveUsers();renderPengguna();}
function hapusSatu(id,tgl){if(!confirm('Hapus?'))return;presensi=presensi.filter(p=>!(p.id==id&&p.tanggal==tgl));save();renderHariIni();renderLaporan();}
function hapusSatuSantri(id){if(currentUser.role!=='admin')return alert('Hanya admin');if(!confirm('Hapus?'))return;santri=santri.filter(s=>s.id!=id);presensi=presensi.filter(p=>p.id!=id);save();renderMaster();renderHariIni();renderLaporan();}
function hapusHariIni(){let t=new Date().toISOString().slice(0,10);let j=presensi.filter(p=>p.tanggal==t).length;if(j==0)return alert('Kosong');if(!confirm('Hapus '+j+' hari ini?'))return;presensi=presensi.filter(p=>p.tanggal!=t);save();renderHariIni();renderLaporan();}
function hapusSemuaPresensi(){if(currentUser.role!=='admin')return alert('Hanya admin');if(!confirm('Hapus semua?'))return;if(!confirm('Yakin?'))return;presensi=[];save();renderHariIni();renderLaporan();}
function hapusSemuaSantri(){if(currentUser.role!=='admin')return alert('Hanya admin');if(!confirm('Hapus semua santri?'))return;if(!confirm('Yakin?'))return;santri=[];presensi=[];save();renderMaster();renderHariIni();renderLaporan();}
function exportExcel(){let t=document.getElementById('filterTanggal').value;let d=t?presensi.filter(p=>p.tanggal==t):presensi;if(d.length==0)return alert('Kosong');let ws=XLSX.utils.json_to_sheet(d);let wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Laporan');XLSX.writeFile(wb,'Laporan.xlsx');}
checkLogin();renderMaster();renderHariIni();renderLaporan();renderPengguna();
