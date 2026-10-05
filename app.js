let santri=JSON.parse(localStorage.getItem('ttq_santri')||'[]');
let presensi=JSON.parse(localStorage.getItem('ttq_presensi')||'[]');
function save(){localStorage.setItem('ttq_santri',JSON.stringify(santri));localStorage.setItem('ttq_presensi',JSON.stringify(presensi));}
function openTab(n){
document.querySelectorAll('.tab').forEach((t,i)=>{t.classList.remove('active');if((n=='presensi'&&i==0)||(n=='master'&&i==1)||(n=='laporan'&&i==2))t.classList.add('active');});
document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
document.getElementById('tab-'+n).classList.add('active');
if(n=='laporan')renderLaporan();if(n=='master')renderMaster();if(n=='presensi')renderHariIni();
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
if(!s){res.innerHTML='❌ IDPPS '+id+' tidak ditemukan';e.target.value='';return;}
let today=new Date().toISOString().slice(0,10);
let sudah=presensi.find(p=>p.id==id&&p.tanggal==today);
if(sudah){res.innerHTML='⚠️ '+s.nama+' sudah hadir jam '+sudah.jam;}
else{
let jam=new Date().toLocaleTimeString('id-ID');
presensi.push({id,nama:s.nama,tanggal:today,jam});save();
res.innerHTML='✅ '+s.nama+' - HADIR jam '+jam;
}
e.target.value='';renderHariIni();renderLaporan();
}
});
function renderHariIni(){
let today=new Date().toISOString().slice(0,10);
let data=presensi.filter(p=>p.tanggal==today).reverse();
document.getElementById('jumlahHadir').innerText=data.length+' Hadir';
document.getElementById('tabelHadirHariIni').innerHTML=data.map(p=>'<tr><td>'+p.jam+'</td><td>'+p.id+'</td><td>'+p.nama+'</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatu(\''+p.id+'\',\''+p.tanggal+'\')">🗑️ Hapus</button></td></tr>').join('')||'<tr><td colspan=4 style="text-align:center">Belum ada</td></tr>';
}
function renderMaster(){
let q=(document.getElementById('cariSantri').value||'').toLowerCase();
let filtered=santri.filter(s=>s.id.toLowerCase().includes(q)||s.nama.toLowerCase().includes(q));
document.getElementById('totalSantri').innerText=santri.length;
document.getElementById('tabelMaster').innerHTML=filtered.map(s=>'<tr><td>'+s.id+'</td><td>'+s.nama+'</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatuSantri(\''+s.id+'\')">🗑️ Hapus</button></td></tr>').join('');
}
function renderLaporan(){
let t=document.getElementById('filterTanggal').value;
let data=t?presensi.filter(p=>p.tanggal==t):presensi;
document.getElementById('tabelLaporan').innerHTML=data.slice().reverse().map(p=>'<tr><td>'+p.tanggal+'</td><td>'+p.id+'</td><td>'+p.nama+'</td><td>'+p.jam+'</td><td><button class="btn btn-danger btn-sm" onclick="hapusSatu(\''+p.id+'\',\''+p.tanggal+'\')">🗑️</button></td></tr>').join('')||'<tr><td colspan=5 style="text-align:center">Tidak ada data</td></tr>';
}
document.getElementById('filterTanggal').addEventListener('change',renderLaporan);
function hapusSatu(id,tanggal){
if(!confirm('Hapus presensi '+id+' tanggal '+tanggal+'?'))return;
presensi=presensi.filter(p=>!(p.id==id&&p.tanggal==tanggal));
save();renderHariIni();renderLaporan();
}
function hapusSatuSantri(id){
if(!confirm('Hapus santri IDPPS '+id+'?'))return;
santri=santri.filter(s=>s.id!=id);
presensi=presensi.filter(p=>p.id!=id);
save();renderMaster();renderHariIni();renderLaporan();
}
function hapusHariIni(){
let today=new Date().toISOString().slice(0,10);
let jml=presensi.filter(p=>p.tanggal==today).length;
if(jml==0)return alert('Tidak ada data hari ini');
if(!confirm('Hapus '+jml+' presensi hari ini ('+today+')? Untuk uji coba. Data santri aman.'))return;
presensi=presensi.filter(p=>p.tanggal!=today);
save();renderHariIni();renderLaporan();alert('Presensi hari ini dihapus!');
}
function hapusTanggalTerpilih(){
let t=document.getElementById('filterTanggal').value;
if(!t)return alert('Pilih tanggal dulu');
let jml=presensi.filter(p=>p.tanggal==t).length;
if(jml==0)return alert('Tidak ada data');
if(!confirm('Hapus '+jml+' data tanggal '+t+'?'))return;
presensi=presensi.filter(p=>p.tanggal!=t);
save();renderHariIni();renderLaporan();alert('Data '+t+' dihapus!');
}
function hapusSemuaPresensi(){
if(presensi.length==0)return alert('Sudah kosong');
if(!confirm('YAKIN hapus SEMUA '+presensi.length+' riwayat?'))return;
if(!confirm('Konfirmasi kedua: benar-benar hapus semua?'))return;
presensi=[];save();renderHariIni();renderLaporan();alert('Semua presensi dihapus!');
}
function hapusSemuaSantri(){
if(santri.length==0)return alert('Sudah kosong');
if(!confirm('YAKIN hapus SEMUA '+santri.length+' santri?'))return;
if(!confirm('Konfirmasi kedua: hapus permanen?'))return;
santri=[];presensi=[];save();renderMaster();renderHariIni();renderLaporan();alert('Semua dihapus!');
}
function exportExcel(){
let t=document.getElementById('filterTanggal').value;
let data=t?presensi.filter(p=>p.tanggal==t):presensi;
if(data.length==0)return alert('Tidak ada data');
let ws=XLSX.utils.json_to_sheet(data);
let wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Laporan');
XLSX.writeFile(wb,'Laporan_TTQ_'+(t||'semua')+'.xlsx');
}
renderMaster();renderHariIni();renderLaporan();
