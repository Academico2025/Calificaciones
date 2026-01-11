
/* =========================
   MODELO DE DATOS (localStorage)
   ========================= */
const ACTIVE_KEY_STORE = "CAL_APP_ACTIVE_IE_LS_KEY";
const KEY_REGISTRY_STORE = "CAL_APP_KNOWN_IE_KEYS";
const DEFAULT_LS_KEY = "CAL_APP_DB_V4_PERIOD_CFG_NOJSON";

function sanitizeLSKey(k){
  k = String(k||"").trim();
  if(!k) return "";
  // Permite letras, números, guion, guion bajo y punto. Reemplaza espacios por "_".
  k = k.replace(/\s+/g,"_").replace(/[^A-Za-z0-9._-]/g,"");
  return k;
}
function escapeAttr(s){
  // escape seguro para atributos HTML
  return escapeHTML(String(s||"")).replace(/"/g,"&quot;");
}
function getActiveLSKey(){
  const k = sanitizeLSKey(localStorage.getItem(ACTIVE_KEY_STORE));
  return k || DEFAULT_LS_KEY;
}
function rememberLSKey(k){
  const kk = sanitizeLSKey(k);
  if(!kk) return;
  try{
    const raw = localStorage.getItem(KEY_REGISTRY_STORE);
    const arr = raw ? JSON.parse(raw) : [];
    const out = Array.isArray(arr) ? arr : [];
    if(!out.includes(kk)) out.unshift(kk);
    localStorage.setItem(KEY_REGISTRY_STORE, JSON.stringify(out.slice(0, 30)));
  }catch(e){}
}
function listKnownLSKeys(){
  const keys = new Set();
  keys.add(DEFAULT_LS_KEY);
  try{
    const raw = localStorage.getItem(KEY_REGISTRY_STORE);
    const arr = raw ? JSON.parse(raw) : [];
    if(Array.isArray(arr)) arr.forEach(k=>keys.add(sanitizeLSKey(k)));
  }catch(e){}
  for(let i=0;i<localStorage.length;i++){
    const k = localStorage.key(i);
    if(!k) continue;
    if(k === DEFAULT_LS_KEY) keys.add(k);
    if(/^CAL_APP_DB_/i.test(k)) keys.add(sanitizeLSKey(k));
  }
  return Array.from(keys).filter(Boolean);
}
function setActiveLSKey(k){
  const kk = sanitizeLSKey(k);
  if(!kk) return;
  localStorage.setItem(ACTIVE_KEY_STORE, kk);
  rememberLSKey(kk);
  location.reload();
}
function getIEFriendlyNameForKey(k){
  try{
    const v = localStorage.getItem("CAL_APP_IE_NAME__"+k);
    return String(v||"").trim();
  }catch(e){ return ""; }
}
function updateHeaderIEBadge(){
  const el = document.getElementById("activeIEBadge");
  if(!el) return;
  const nm = getIEFriendlyNameForKey(LS_KEY);
  el.innerHTML = nm
    ? `IE: <strong>${escapeHTML(nm)}</strong> <span class="mono">(${escapeHTML(LS_KEY)})</span>`
    : `IE: <strong class="mono">${escapeHTML(LS_KEY)}</strong>`;
}

let LS_KEY = getActiveLSKey();
rememberLSKey(LS_KEY);

function openIESwitcher(){
  const modal = document.getElementById("ieModal");
  if(!modal) return;

  const cur = document.getElementById("ieCurrentKey");
  if(cur) cur.textContent = LS_KEY;

  const sel = document.getElementById("ieKeySelect");
  const inp = document.getElementById("ieKeyInput");
  const name = document.getElementById("ieKeyName");

  if(sel){
    const opts = listKnownLSKeys();
    sel.innerHTML = opts.map(k=>`<option value="${escapeAttr(k)}">${escapeHTML(k)}</option>`).join("");
    sel.value = LS_KEY;
    sel.onchange = ()=>{ if(inp) inp.value = sel.value; };
  }
  if(inp) inp.value = LS_KEY;
  if(name) name.value = getIEFriendlyNameForKey(LS_KEY);

  modal.classList.remove("hide");
}
function closeIESwitcher(){
  const modal = document.getElementById("ieModal");
  if(modal) modal.classList.add("hide");
}
function applyIESwitcher(){
  const inp = document.getElementById("ieKeyInput");
  const name = document.getElementById("ieKeyName");
  const v = sanitizeLSKey(inp ? inp.value : "");
  if(!v){
    alert("Ingresa un código de IE (LS_KEY) válido.");
    return;
  }
  if(name){
    const nm = String(name.value||"").trim();
    if(nm){
      try{ localStorage.setItem("CAL_APP_IE_NAME__"+v, nm); }catch(e){}
    }
  }
  setActiveLSKey(v);
}
function deleteCurrentIEData(){
  if(!confirm("Esto borrará TODOS los datos (estudiantes, áreas, notas y configuración) de esta IE (solo de este navegador). ¿Continuar?")) return;
  localStorage.removeItem(LS_KEY);
  location.reload();
}
function uid(prefix="id"){ return prefix + "_" + Math.random().toString(16).slice(2) + "_" + Date.now().toString(16); }
function nowISO(){ return new Date().toISOString(); }
function clamp(n, a, b){ return Math.max(a, Math.min(b, n)); }
function toNum(v){
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(String(v).replace(",", ".").trim());
  return Number.isFinite(n) ? n : null;
}
function isNum(x){ return typeof x === "number" && Number.isFinite(x); }
function fmt(n, d=2){ return isNum(n) ? n.toFixed(d).replace(".", ",") : ""; } // coma

function defaultPeriodCfg(periods){
  const cfg = {};
  for(let p=1;p<=periods;p++){
    cfg[p] = { wPr:0.50, wCo:0.40, wAc:0.10, prNotes:3, coNotes:3 };
  }
  return cfg;
}

function defaultPeriodDates(periods){
  const out = {};
  for(let p=1;p<=periods;p++){
    out[p] = { start:"", end:"" };
  }
  return out;
}
function normalizePeriodDates(pd, periods){
  const out = {};
  for(let p=1;p<=periods;p++){
    const src = (pd && pd[p]) ? pd[p] : {};
    out[p] = {
      start: String(src.start || "").trim(),
      end: String(src.end || "").trim()
    };
  }
  return out;
}
function fmtISOToDMY(iso){
  const s = String(iso || "").trim();
  if(!s) return "";
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if(!m) return s; // fallback si no es ISO
  return `${m[3]}/${m[2]}/${m[1]}`;
}

function setPeriodDateField(p, field, val){
  const periods = db.settings.periods || 3;
  const pp = Number(p);
  if(!pp || pp < 1 || pp > periods) return;
  if(field !== "start" && field !== "end") return;

  if(!db.settings.periodDates) db.settings.periodDates = defaultPeriodDates(periods);
  db.settings.periodDates = normalizePeriodDates(db.settings.periodDates, periods);

  db.settings.periodDates[pp][field] = String(val || "").trim();
  saveDB();
}

function renderPeriodDatesTable(){
  const tb = document.getElementById("periodDatesTbody");
  if(!tb) return;
  const periods = db.settings.periods || 3;
  if(!db.settings.periodDates) db.settings.periodDates = defaultPeriodDates(periods);
  db.settings.periodDates = normalizePeriodDates(db.settings.periodDates, periods);

  tb.innerHTML = Array.from({length:periods}, (_,i)=>{
    const p = i+1;
    const pd = db.settings.periodDates[p] || {start:"", end:""};
    return `
      <tr>
        <td class="center mono">${p}</td>
        <td><input type="date" value="${escapeHTML(pd.start||"")}" onchange="setPeriodDateField(${p},'start',this.value)"></td>
        <td><input type="date" value="${escapeHTML(pd.end||"")}" onchange="setPeriodDateField(${p},'end',this.value)"></td>
      </tr>
    `;
  }).join("");
}


function defaultDB(){
  const today = new Date();
  const dd = String(today.getDate()).padStart(2,"0");
  const mm = String(today.getMonth()+1).padStart(2,"0");
  const yyyy = String(today.getFullYear());
  const periods = 3;
  return {
    settings: {
      periods,
      periodCfg: defaultPeriodCfg(periods),
      periodDates: defaultPeriodDates(periods),
      gradeGlobal:"",
      institution:"",
      place:"",
      municipio:"",
      depto:"",
      // Boletín
      jornada:"MAÑANA",
      year: yyyy,
      grupo:"",
      fechaBoletin: `${dd}/${mm}/${yyyy}`,
      promovido:"SI",
      alGrado:"",
      actMej:"",
      obsBoletin:"",
      // Encabezado/Firmas/Sellos
      ieName:"",
      nit:"",
      dane:"",
      rectorName:"",
      dirName:"",
      images: {
        shield:"",
        rectorSign:"",
        rectorSeal:"",
        dirSign:"",
        dirSeal:""
      }
    },
    students: [],
    areas: [],
    grades: {}, // areaId -> period -> studentId -> row
    updatedAt: nowISO()
  };
}
let db = loadDB();

migrateBoletinPerStudent();
function loadDB(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if(!raw) return defaultDB();
    const parsed = JSON.parse(raw);
    const def = defaultDB();
    parsed.settings = Object.assign(def.settings, parsed.settings || {});
    if(!parsed.settings.images) parsed.settings.images = def.settings.images;
    parsed.settings.images = Object.assign(def.settings.images, parsed.settings.images || {});
    if(!parsed.settings.periodCfg) parsed.settings.periodCfg = defaultPeriodCfg(parsed.settings.periods || 3);
    parsed.settings.periodCfg = normalizePeriodCfg(parsed.settings.periodCfg, parsed.settings.periods || 3);
    
    if(!parsed.settings.periodDates) parsed.settings.periodDates = defaultPeriodDates(parsed.settings.periods || 3);
    parsed.settings.periodDates = normalizePeriodDates(parsed.settings.periodDates, parsed.settings.periods || 3);
if(!parsed.students) parsed.students = [];
    if(!parsed.areas) parsed.areas = [];
    if(!parsed.grades) parsed.grades = {};
    if(!parsed.updatedAt) parsed.updatedAt = nowISO();
    return parsed;
  }catch(e){
    console.warn("DB corrupta, restaurando defaults", e);
    return defaultDB();
  }
}
function isQuotaExceeded(err){
  if(!err) return false;
  const name = String(err.name||"").toLowerCase();
  const msg  = String(err.message||"").toLowerCase();
  return name.includes("quota") || msg.includes("quota") || msg.includes("exceeded") || msg.includes("storage");
}

function saveDB(){
  db.updatedAt = nowISO();
  try{
    localStorage.setItem(LS_KEY, JSON.stringify(db));
  }catch(e){
    console.error("No se pudo guardar en localStorage", e);
    if(isQuotaExceeded(e)){
      alert(
`No se pudo guardar porque el almacenamiento del navegador está lleno (cuota de localStorage).

Recomendaciones:
• Exporta un Respaldo JSON desde Configuración.
• Reduce el tamaño de imágenes (escudo/firmas/sellos).
• Usa un LS_KEY diferente (otro perfil) si es necesario.`
);
}else{
      alert("No se pudo guardar la información localmente. Revisa la consola para más detalle.");
    }
    // No hacemos throw: evitamos romper la UI, pero la app debe considerar que el cambio NO persistió.
  }
  refreshAllMeta();
}



function exportBackupJSON(){
  try{
    const metaName = (db.settings?.ieName || db.settings?.institution || "IE").trim() || "IE";
    const stamp = nowISO().replace(/[:.]/g,"-");
    const filename = safeFileName(`RESPALDO_${metaName}_${stamp}.json`);
    const blob = new Blob([JSON.stringify(db, null, 2)], {type:"application/json"});
    downloadBlob(blob, filename);
    toast("Respaldo JSON exportado.");
  }catch(e){
    console.error(e);
    alert("No se pudo exportar el respaldo JSON.");
  }
}

function openImportBackupJSON(){
  const el = document.getElementById("backupJsonFile");
  if(!el){ alert("No se encontró el selector de archivo JSON."); return; }
  el.value = "";
  el.click();
}

async function importBackupJSON(file){
  if(!file) return;
  try{
    const txt = await file.text();
    const parsed = JSON.parse(txt);
    if(!confirm("Vas a RESTAURAR un respaldo JSON.\n\nEsto reemplazará la información de la base actual (LS_KEY activo).\n¿Deseas continuar?")) return;

    // Guardar tal cual y luego recargar para que loadDB() normalice estructura y aplique defaults.
    localStorage.setItem(LS_KEY, JSON.stringify(parsed));
    toast("Respaldo restaurado. Recargando…");
    setTimeout(()=> location.reload(), 250);
  }catch(e){
    console.error(e);
    alert("No se pudo restaurar el respaldo. Verifica que el archivo sea un JSON válido.");
  }
}

function migrateBoletinPerStudent(){
  // Migra 'Promovido / Al grado' para que sea por estudiante (SI / NO / ND).
  // Se ejecuta una sola vez por base de datos (localStorage).
  try{
    db._migrations = db._migrations || {};
    if(db._migrations.promovidoPerStudent) return;

    const year = String(db.settings?.year || "").trim();
    const defProm = year ? String(db.settings?.promovido || "ND").trim() : "";
    const defAl = String(db.settings?.alGrado || "").trim();

    (db.students || []).forEach(s=>{
      if(!s) return;
      if(!s.boletin) s.boletin = {};
      if(s.boletin.promovido === undefined || s.boletin.promovido === null){
        s.boletin.promovido = defProm;
      }
      if(s.boletin.alGrado === undefined || s.boletin.alGrado === null){
        s.boletin.alGrado = defAl;
      }
    });

    db._migrations.promovidoPerStudent = true;
    // Guardar sin forzar render (puede ejecutarse antes de que exista la UI)
    db.updatedAt = nowISO();
    localStorage.setItem(LS_KEY, JSON.stringify(db));
  }catch(e){
    console.warn("No se pudo migrar promovido por estudiante", e);
  }
}

function getBoletinPromovidoForStudent(st){
  const year = String(db.settings?.year || "").trim();
  if(!year) return "";
  const v = String(st?.boletin?.promovido ?? "").trim().toUpperCase();
  return v || "ND";
}
function getBoletinAlGradoForStudent(st){
  return String(st?.boletin?.alGrado ?? "").trim();
}

function normalizePeriodCfg(pc, periods){
  const out = {};
  const def = defaultPeriodCfg(periods);
  for(let p=1;p<=periods;p++){
    const src = pc?.[p] || {};
    out[p] = {
      wPr: isNum(toNum(src.wPr)) ? clamp(toNum(src.wPr),0,1) : def[p].wPr,
      wCo: isNum(toNum(src.wCo)) ? clamp(toNum(src.wCo),0,1) : def[p].wCo,
      wAc: isNum(toNum(src.wAc)) ? clamp(toNum(src.wAc),0,1) : def[p].wAc,
      prNotes: Math.max(1, Math.min(6, Math.round(toNum(src.prNotes) ?? def[p].prNotes))),
      coNotes: Math.max(1, Math.min(6, Math.round(toNum(src.coNotes) ?? def[p].coNotes))),
    };
  }
  return out;
}
function getPeriodCfg(period){
  const p = Number(period);
  const periods = db.settings.periods;
  if(!db.settings.periodCfg) db.settings.periodCfg = defaultPeriodCfg(periods);
  if(!db.settings.periodCfg[p]) db.settings.periodCfg = normalizePeriodCfg(db.settings.periodCfg, periods);
  return db.settings.periodCfg[p] || defaultPeriodCfg(periods)[p];
}

/* =========================
   ESCALAS (según plantillas)
   ========================= */
function escalaNacional_D(nota){
  if(!isNum(nota)) return " ";
  if(nota <= 2.99) return "D. BAJO";
  if(nota <= 3.99) return "D. BASICO";
  if(nota <= 4.59) return "D. ALTO";
  return "D. SUPERIOR";
}
function escalaNacional_SIN_D(nota){
  if(!isNum(nota)) return " ";
  if(nota <= 2.99) return "BAJO";
  if(nota <= 3.99) return "BASICO";
  if(nota <= 4.59) return "ALTO";
  return "SUPERIOR";
}

/* =========================
   Excel-like: AVERAGEIF(">0,9")
   ========================= */
function avgIfGT09(arr){
  const nums = arr.map(toNum).filter(v => isNum(v) && v > 0.9);
  if(!nums.length) return null;
  const avg = nums.reduce((a,b)=>a+b,0)/nums.length;
  return clamp(avg,0,5);
}
function avgIfGT09_periodNotes(arr){
  const nums = arr.filter(v => isNum(v) && v > 0.9);
  if(!nums.length) return null;
  return clamp(nums.reduce((a,b)=>a+b,0)/nums.length,0,5);
}
function avgSimple(arr){
  const nums = arr.filter(v => isNum(v));
  if(!nums.length) return null;
  return clamp(nums.reduce((a,b)=>a+b,0)/nums.length,0,5);
}

/* =========================
   FILA (estructura AREAS.xlsx) + compat (hasta 6 notas)
   ========================= */
function emptyRow(){
  return {
    j: 0, sj: 0,
    pr_n1:null, pr_n2:null, pr_n3:null, pr_n4:null, pr_n5:null, pr_n6:null,
    pr_ae:null, pr_st:null, pr_rec:null,
    co_n1:null, co_n2:null, co_n3:null, co_n4:null, co_n5:null, co_n6:null,
    co_ae:null, co_st:null, co_rec:null,
    ac_doc:null, ac_est:null, ac_st:null, ac_rec:null,
    _computed: { pr_st:null, pr_w:null, co_st:null, co_w:null, ac_st:null, ac_w:null, nota:null }
  };
}
function noteField(prefix, idx){ return `${prefix}_n${idx}`; }
function getNotesArray(row, prefix, count){
  const out = [];
  for(let i=1;i<=count;i++){
    out.push(row[noteField(prefix,i)]);
  }
  return out;
}

/* =========================
   CÁLCULO IDÉNTICO A EXCEL (pero con config por periodo)
   ========================= */
function calcRowExcelLike(row, cfg){
  const j = toNum(row.j) ?? 0;
  const sj = toNum(row.sj) ?? 0;
  row.j = Math.max(0, Math.round(j));
  row.sj = Math.max(0, Math.round(sj));

  const prCount = cfg?.prNotes ?? 3;
  const coCount = cfg?.coNotes ?? 3;

  const pr_st_auto = avgIfGT09([...getNotesArray(row,"pr",prCount), row.pr_ae, row.pr_rec]);
  let pr_st = toNum(row.pr_st);
  pr_st = isNum(pr_st) ? clamp(pr_st,0,5) : pr_st_auto;
  const pr_w = isNum(pr_st) ? pr_st * (cfg?.wPr ?? 0.5) : null;

  const co_st_auto = avgIfGT09([...getNotesArray(row,"co",coCount), row.co_ae, row.co_rec]);
  let co_st = toNum(row.co_st);
  co_st = isNum(co_st) ? clamp(co_st,0,5) : co_st_auto;
  const co_w = isNum(co_st) ? co_st * (cfg?.wCo ?? 0.4) : null;

  const ac_st_auto = avgIfGT09([row.ac_doc,row.ac_est, row.ac_rec]);
  let ac_st = toNum(row.ac_st);
  ac_st = isNum(ac_st) ? clamp(ac_st,0,5) : ac_st_auto;
  const ac_w = isNum(ac_st) ? ac_st * (cfg?.wAc ?? 0.1) : null;

  let nota = null;
  const parts = [pr_w, co_w, ac_w].filter(isNum);
  if(parts.length) nota = clamp(parts.reduce((a,b)=>a+b,0),0,5);

  row._computed = { pr_st, pr_w, co_st, co_w, ac_st, ac_w, nota };
  return row;
}

/* =========================
   NAV / UI
   ========================= */
function refreshAllMeta(){
  document.getElementById("kStudents").textContent = db.students.length;
  document.getElementById("kAreas").textContent = db.areas.length;
  document.getElementById("kPeriods").textContent = db.settings.periods;
  document.getElementById("kUpdated").textContent = new Date(db.updatedAt).toLocaleString();

  let count=0;
  for(const aId in db.grades){
    const byP = db.grades[aId] || {};
    for(const p in byP) count += Object.keys(byP[p]||{}).length;
  }
  document.getElementById("kEntries").textContent = count;

  const setPeriods = document.getElementById("setPeriods");
  if(setPeriods) setPeriods.value = String(db.settings.periods);

  const gg = document.getElementById("setGradeGlobal"); if(gg) gg.value = db.settings.gradeGlobal || "";
  const inst = document.getElementById("setInstitution"); if(inst) inst.value = db.settings.institution || "";
  const mun = document.getElementById("setMunicipio"); if(mun) mun.value = db.settings.municipio || "";
  const dep = document.getElementById("setDepto"); if(dep) dep.value = db.settings.depto || "";
  const place = document.getElementById("setPlace"); if(place) place.value = db.settings.place || "";

  // Boletín
  const jor = document.getElementById("setJornada"); if(jor) jor.value = db.settings.jornada || "MAÑANA";
  const year = document.getElementById("setYear"); if(year) year.value = db.settings.year || "";
  const grp = document.getElementById("setGrupo"); if(grp) grp.value = db.settings.grupo || "";
  const fbo = document.getElementById("setFechaBoletin"); if(fbo) fbo.value = db.settings.fechaBoletin || "";
  const sid = document.getElementById("genStudent")?.value;
  const stSel = sid ? db.students.find(s=>s.id===sid) : null;
  const yearHas = String(db.settings.year||"").trim();
  const pro = document.getElementById("setPromovido");
  if(pro) pro.value = yearHas ? getBoletinPromovidoForStudent(stSel) : "";
  const alg = document.getElementById("setAlGrado");
  if(alg) alg.value = getBoletinAlGradoForStudent(stSel) || "";
  const acm = document.getElementById("setActMej"); if(acm) acm.value = db.settings.actMej || "";
  const obs = document.getElementById("setObsBoletin"); if(obs) obs.value = db.settings.obsBoletin || "";

  // Encabezado/firmas/sellos
  const ien = document.getElementById("setIEName"); if(ien) ien.value = db.settings.ieName || "";
  const nit = document.getElementById("setNIT"); if(nit) nit.value = db.settings.nit || "";
  const dane = document.getElementById("setDANE"); if(dane) dane.value = db.settings.dane || "";
  const rn = document.getElementById("setRectorName"); if(rn) rn.value = db.settings.rectorName || "";
  const dn = document.getElementById("setDirName"); if(dn) dn.value = db.settings.dirName || "";

  // previews
  setPreview("prevShield", db.settings.images?.shield);
  setPreview("prevRectorSign", db.settings.images?.rectorSign);
  setPreview("prevRectorSeal", db.settings.images?.rectorSeal);
  setPreview("prevDirSign", db.settings.images?.dirSign);
  setPreview("prevDirSeal", db.settings.images?.dirSeal);

  renderPeriodCfgTable();
  renderPeriodDatesTable();
  renderPeriodDatesTable();
}
function setPreview(imgId, dataUrl){
  const el = document.getElementById(imgId);
  if(!el) return;
  if(dataUrl){
    el.src = dataUrl;
    el.style.display = "block";
  }else{
    el.removeAttribute("src");
    el.style.display = "none";
  }
}

function go(view){
  document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("active", t.dataset.view===view));
  document.querySelectorAll("main section").forEach(sec=>sec.classList.add("hide"));
  document.getElementById("view-"+view).classList.remove("hide");

  if(view==="students") renderStudents();
  if(view==="areas") renderAreas();
  if(view==="gradebook") { fillAreaSelects(); fillPeriodSelects(); renderGradebook(); }
  if(view==="areaFinal") { fillAreaSelects(); renderFinalArea(); }
  if(view==="generator") { fillStudentSelect(); fillPeriodSelectsGenerator(); renderGenerator(); refreshAllMeta(); }
  if(view==="certificates") { fillCertStudentSelect(); fillCertResolution(); renderCertificate(); refreshAllMeta(); }
  if(view==="settings") refreshAllMeta();
  if(view==="dashboard") refreshAllMeta();
}
document.querySelectorAll(".tab").forEach(btn=>{ const v = btn.dataset.view; if(!v) return; btn.addEventListener("click", ()=> go(v)); });

function saveBoletinSettings(){
  db.settings.jornada = (document.getElementById("setJornada").value||"MAÑANA").trim();

  const _year = (document.getElementById("setYear").value||"").trim();
  db.settings.year = _year;

  db.settings.grupo = (document.getElementById("setGrupo").value||"").trim();
  db.settings.fechaBoletin = (document.getElementById("setFechaBoletin").value||"").trim();

const _proEl = document.getElementById("setPromovido");
let _pro = (_proEl?.value||"").trim().toUpperCase();
if(!_year){
  // Si no hay año lectivo, no se muestra en el boletín, pero NO borra el valor guardado por estudiante
  if(_proEl) _proEl.value = "";
}else{
  if(!_pro) _pro = "ND";
  if(_proEl) _proEl.value = _pro;
}

const _al = (document.getElementById("setAlGrado").value||"").trim();

// Guardar por estudiante (no global)
const sid = document.getElementById("genStudent")?.value;
const stSel = sid ? db.students.find(s=>s.id===sid) : null;
if(stSel){
  if(!stSel.boletin) stSel.boletin = {};
  if(_year){
    stSel.boletin.promovido = _pro || "ND";
    stSel.boletin.alGrado = _al;
  }
}

// Mantener valores en settings como "default" (compatibilidad / nuevos estudiantes)
db.settings.promovido = (_year ? (_pro || "ND") : "");
db.settings.alGrado = _al;
  db.settings.actMej = (document.getElementById("setActMej").value||"").trim();
  db.settings.obsBoletin = (document.getElementById("setObsBoletin").value||"").trim();
  saveDB();
  renderGenerator();
}

/* =========================
   IMÁGENES (escudo / firmas / sellos)
   ========================= */
async function handleImageUpload(evt, key){
  const file = evt.target.files?.[0];
  evt.target.value = "";
  if(!file) return;

  if(!/^image\/(png|jpeg)$/.test(file.type)){
    alert("Solo se permite JPG o PNG.");
    return;
  }

  // Permitimos archivos más grandes porque la app los optimiza antes de guardar.
  if(file.size > 10_000_000){
    alert("La imagen supera 10MB. Por favor usa una versión más liviana.");
    return;
  }

  try{
    const dataUrl = await processImageForStorage(file, key);
    db.settings.images[key] = dataUrl;
    saveDB();
    refreshAllMeta();
    renderGenerator();
  }catch(e){
    console.error(e);
    alert("No se pudo cargar/optimizar la imagen.");
  }
}
function clearImage(key){
  if(!confirm("¿Quitar esta imagen?")) return;
  db.settings.images[key] = "";
  saveDB();
  refreshAllMeta();
  renderGenerator();
}
function fileToDataURL(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onload = ()=> resolve(String(reader.result||""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function dataUrlApproxBytes(dataUrl){
  try{
    const comma = String(dataUrl||"").indexOf(",");
    const b64 = comma>=0 ? dataUrl.slice(comma+1) : String(dataUrl||"");
    // aprox: 3/4 del base64, menos padding
    const pad = (b64.endsWith("==") ? 2 : (b64.endsWith("=") ? 1 : 0));
    return Math.max(0, Math.floor((b64.length * 3) / 4) - pad);
  }catch{ return 0; }
}
function keyImageProfile(key){
  switch(String(key||"")){
    case "shield":      return { maxDim: 900, outType: "image/jpeg", quality: 0.85, targetBytes: 900_000 };
    case "rectorSign":  return { maxDim: 800, outType: "image/png",  quality: 0.92, targetBytes: 350_000 };
    case "dirSign":     return { maxDim: 800, outType: "image/png",  quality: 0.92, targetBytes: 350_000 };
    case "rectorSeal":  return { maxDim: 800, outType: "image/png",  quality: 0.92, targetBytes: 500_000 };
    case "dirSeal":     return { maxDim: 800, outType: "image/png",  quality: 0.92, targetBytes: 500_000 };
    default:            return { maxDim: 900, outType: "image/jpeg", quality: 0.85, targetBytes: 900_000 };
  }
}
function loadImageFromDataUrl(dataUrl){
  return new Promise((resolve, reject)=>{
    const img = new Image();
    img.onload = ()=> resolve(img);
    img.onerror = reject;
    img.src = dataUrl;
  });
}
async function compressDataUrl(dataUrl, profile){
  const img = await loadImageFromDataUrl(dataUrl);
  let {maxDim, outType, quality, targetBytes} = profile;

  // si el usuario sube PNG con transparencia pero vamos a JPEG, fondo blanco
  const needsWhiteBg = (outType === "image/jpeg");

  let w = img.naturalWidth || img.width;
  let h = img.naturalHeight || img.height;
  if(!w || !h) return dataUrl;

  let scale = Math.min(1, maxDim / Math.max(w,h));
  let curDataUrl = dataUrl;

  for(let attempt=0; attempt<5; attempt++){
    const tw = Math.max(1, Math.round(w * scale));
    const th = Math.max(1, Math.round(h * scale));
    const canvas = document.createElement("canvas");
    canvas.width = tw; canvas.height = th;
    const ctx = canvas.getContext("2d");

    if(needsWhiteBg){
      ctx.fillStyle = "#fff";
      ctx.fillRect(0,0,tw,th);
    }
    ctx.drawImage(img, 0, 0, tw, th);

    let out = "";
    if(outType === "image/jpeg"){
      out = canvas.toDataURL("image/jpeg", quality);
    }else{
      out = canvas.toDataURL("image/png");
    }

    curDataUrl = out;
    const bytes = dataUrlApproxBytes(out);

    // si ya estamos dentro del objetivo (o el objetivo no aplica), terminamos
    if(!targetBytes || bytes <= targetBytes){
      return out;
    }

    // Estrategia de reducción: si es JPEG, baja calidad; si es PNG o sigue grande, baja dimensiones.
    if(outType === "image/jpeg" && quality > 0.55){
      quality = Math.max(0.55, quality - 0.08);
    }else{
      scale = scale * 0.85;
      if(scale < 0.15) return out;
    }
  }
  return curDataUrl;
}
async function processImageForStorage(file, key){
  const dataUrl = await fileToDataURL(file);
  const profile = keyImageProfile(key);
  return await compressDataUrl(dataUrl, profile);
}

/* =========================
   CONFIG POR PERIODO (UI + persistencia)
   ========================= */
function renderPeriodCfgTable(){
  const tb = document.getElementById("periodCfgTbody");
  if(!tb) return;
  const periods = db.settings.periods;
  db.settings.periodCfg = normalizePeriodCfg(db.settings.periodCfg, periods);

  tb.innerHTML = Array.from({length:periods}, (_,i)=>{
    const p = i+1;
    const cfg = getPeriodCfg(p);
    const wp = Math.round((cfg.wPr||0)*100);
    const wc = Math.round((cfg.wCo||0)*100);
    const wa = Math.round((cfg.wAc||0)*100);
    return `
      <tr>
        <td class="center mono">${p}</td>
        <td><input class="num" value="${escapeHTML(String(wp))}" oninput="setPeriodCfgField(${p},'wPr',this.value)"></td>
        <td><input class="num" value="${escapeHTML(String(wc))}" oninput="setPeriodCfgField(${p},'wCo',this.value)"></td>
        <td><input class="num" value="${escapeHTML(String(wa))}" oninput="setPeriodCfgField(${p},'wAc',this.value)"></td>
        <td><input class="num" value="${escapeHTML(String(cfg.prNotes||3))}" oninput="setPeriodCfgField(${p},'prNotes',this.value)"></td>
        <td><input class="num" value="${escapeHTML(String(cfg.coNotes||3))}" oninput="setPeriodCfgField(${p},'coNotes',this.value)"></td>
        <td class="center">
          <button class="btn ghost" onclick="restorePeriodCfg(${p})">Restaurar</button>
        </td>
      </tr>
    `;
  }).join("");
}
function setPeriodCfgField(period, field, value){
  const periods = db.settings.periods;
  db.settings.periodCfg = normalizePeriodCfg(db.settings.periodCfg, periods);
  const cfg = db.settings.periodCfg[period] || defaultPeriodCfg(periods)[period];

  if(field==="wPr" || field==="wCo" || field==="wAc"){
    const v = toNum(value);
    cfg[field] = isNum(v) ? clamp(v/100, 0, 1) : cfg[field];
  }else if(field==="prNotes" || field==="coNotes"){
    const v = toNum(value);
    cfg[field] = Math.max(1, Math.min(6, Math.round(v ?? cfg[field])));
  }
  db.settings.periodCfg[period] = cfg;
  saveDB();

  const currentP = Number(document.getElementById("gbPeriod")?.value || 0);
  if(currentP === Number(period)){
    renderGradebook();
  }
}
function restorePeriodCfg(period){
  const periods = db.settings.periods;
  const def = defaultPeriodCfg(periods)[period];
  db.settings.periodCfg[period] = {...def};
  saveDB();
  renderPeriodCfgTable();
  const currentP = Number(document.getElementById("gbPeriod")?.value || 0);
  if(currentP === Number(period)) renderGradebook();
}

/* =========================
   CRUD ESTUDIANTES
   ========================= */
function addStudent(){
  const name = (document.getElementById("stName").value||"").trim();
  const ident = (document.getElementById("stId").value||"").trim();
  const grade = (document.getElementById("stGrade").value||"").trim();
  const code = (document.getElementById("stCode").value||"").trim();
  if(!name){ alert("Ingresa el nombre del estudiante."); return; }

  db.students.push({ id: uid("st"), name, ident, grade, code });
  saveDB();
  clearStudentForm();
  renderStudents();
  fillStudentSelect();
  renderGradebookRows();
  renderFinalArea();
  renderGenerator();
}
function clearStudentForm(){
  document.getElementById("stName").value="";
  document.getElementById("stId").value="";
  document.getElementById("stGrade").value="";
  document.getElementById("stCode").value="";
}
function sortStudents(){
  db.students.sort((a,b)=> (a.name||"").localeCompare((b.name||""), "es"));
  saveDB();
  renderStudents();
}
function updateStudent(id){
  const st = db.students.find(s=>s.id===id); if(!st) return;
  const name = prompt("APELLIDOS Y NOMBRES:", st.name||""); if(name===null) return;
  const ident = prompt("IDENTIFICACION:", st.ident||""); if(ident===null) return;
  const grade = prompt("GRADO:", st.grade||""); if(grade===null) return;
  const code = prompt("CÓDIGO (opcional):", st.code||""); if(code===null) return;
  st.name = (name||"").trim();
  st.ident = (ident||"").trim();
  st.grade = (grade||"").trim();
  st.code = (code||"").trim();
  saveDB();
  renderStudents();
  fillStudentSelect();
  renderGradebookRows();
  renderFinalArea();
  renderGenerator();
}
function deleteStudent(id){
  const st = db.students.find(s=>s.id===id); if(!st) return;
  if(!confirm(`¿Eliminar estudiante?\n\n${st.name}`)) return;
  db.students = db.students.filter(s=>s.id!==id);

  for(const aId in db.grades){
    const byP = db.grades[aId] || {};
    for(const p in byP){
      if(byP[p] && byP[p][id]) delete byP[p][id];
    }
  }
  saveDB();
  renderStudents(); fillStudentSelect(); renderGradebookRows(); renderFinalArea(); renderGenerator();
}
function renderStudents(){
  const tb = document.getElementById("studentsTbody");
  const q = (document.getElementById("stFilter")?.value||"").toLowerCase().trim();
  let rows = db.students.slice();
  if(q){
    rows = rows.filter(s =>
      (s.name||"").toLowerCase().includes(q) ||
      (s.ident||"").toLowerCase().includes(q) ||
      (s.grade||"").toLowerCase().includes(q) ||
      (s.code||"").toLowerCase().includes(q)
    );
  }
  tb.innerHTML = rows.map((s,i)=>`
    <tr>
      <td class="mono">${i+1}</td>
      <td>${escapeHTML(s.name||"")}</td>
      <td class="mono">${escapeHTML(s.ident||"")}</td>
      <td>${escapeHTML(s.grade||"")}</td>
      <td class="mono">${escapeHTML(s.code||"")}</td>
      <td class="right">
        <button class="btn ghost" onclick="updateStudent('${s.id}')">Editar</button>
        <button class="btn danger" onclick="deleteStudent('${s.id}')">Eliminar</button>
      </td>
    </tr>
  `).join("");
}

/* =========================
   IMPORTAR ESTUDIANTES DESDE XLSX (ACTUALIZA SI YA EXISTE)
   ========================= */
async function importStudentsFromXLSX(evt){
  try{
    if(typeof XLSX === "undefined"){
      alert("No se cargó la librería XLSX. Revisa la conexión (CDN).");
      return;
    }

    const file = evt.target.files?.[0];
    evt.target.value = "";
    if(!file) return;

    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });

    const sheetName = wb.SheetNames?.[0];
    if(!sheetName){
      alert("El archivo no tiene hojas.");
      return;
    }
    const ws = wb.Sheets[sheetName];

    // Leer como matriz (AOA) — soporta celdas vacías
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false, defval: "" });

    if(!aoa || aoa.length < 2){
      alert("El archivo no tiene datos suficientes.");
      return;
    }

    // Normalizador de encabezados (quita tildes, símbolos, dobles espacios)
    const norm = (v) => String(v ?? "")
      .trim()
      .toUpperCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // sin tildes
      .replace(/[º°]/g, "") // N° / Nº
      .replace(/\./g, "")
      .replace(/\s+/g, " ");

    // Buscar fila de encabezados (por si hay filas de título arriba)
    let headerRowIndex = -1;
    for(let i=0; i<Math.min(aoa.length, 10); i++){
      const row = aoa[i] || [];
      const h = row.map(norm);
      if(h.some(x=>x.includes("APELLIDOS")) && h.some(x=>x==="GRADO")){
        headerRowIndex = i;
        break;
      }
    }
    if(headerRowIndex === -1) headerRowIndex = 0;

    const hdrRaw = aoa[headerRowIndex] || [];
    const hdr = hdrRaw.map(norm);

    const idxOf = (...candidates) => {
      for(const cand of candidates){
        const n = norm(cand);
        for(let i=0;i<hdr.length;i++){
          if(hdr[i] === n) return i;
        }
      }
      return -1;
    };

    let IDX_NAME  = idxOf("APELLIDOS Y NOMBRES");
    let IDX_ID    = idxOf("IDENTIFICACION", "DENTIFICACION", "IDENTIFICACIÓN", "DOCUMENTO", "NUMERO DE DOCUMENTO", "NRO DOCUMENTO");
    let IDX_GRADE = idxOf("GRADO", "CURSO");
    let IDX_CODE  = idxOf("CODIGO", "CÓDIGO");

    // Fallback: si el archivo tiene columnas extra vacías, validar SOLO las primeras 5 columnas
    const hdrFirst5 = hdr.slice(0,5);
    const exp = ["N", "APELLIDOS Y NOMBRES", "IDENTIFICACION", "GRADO", "CODIGO"].map(norm);
    const okFirst5 = exp.every((h,i)=> (hdrFirst5[i]||"") === h);

    if((IDX_NAME < 0 || IDX_ID < 0 || IDX_GRADE < 0 || IDX_CODE < 0) && okFirst5){
      IDX_NAME  = 1;
      IDX_ID    = 2;
      IDX_GRADE = 3;
      IDX_CODE  = 4;
    }

    if(IDX_NAME < 0 || IDX_ID < 0 || IDX_GRADE < 0 || IDX_CODE < 0){
      alert(
        "Encabezados inválidos.\n\n" +
        "Se requiere (en este orden):\n" +
        "Nº, APELLIDOS Y NOMBRES, IDENTIFICACION, GRADO, CODIGO\n\n" +
        "Encabezados encontrados:\n" + hdr.filter(Boolean).join(" | ")
      );
      return;
    }

    // Índices existentes para actualización
    const byIdent = new Map();
    const byNameGrade = new Map(); // key -> studentId (para casos sin identificación)
    db.students.forEach(s=>{
      const ident = (s.ident||"").trim();
      if(ident) byIdent.set(ident, s.id);
      const keyNG = (((s.name||"").trim()+"||"+(s.grade||"").trim()).toUpperCase());
      if(keyNG !== "||") byNameGrade.set(keyNG, s.id);
    });

    let created = 0;
    let updated = 0;
    let skippedEmpty = 0;

    for(let r=headerRowIndex+1;r<aoa.length;r++){
      const row = aoa[r] || [];
      const name  = String(row[IDX_NAME] ?? "").trim();
      const ident = String(row[IDX_ID] ?? "").trim();
      const grade = String(row[IDX_GRADE] ?? "").trim();
      const code  = String(row[IDX_CODE] ?? "").trim();

      if(!name){
        skippedEmpty++;
        continue;
      }

      let existingId = null;

      if(ident && byIdent.has(ident)){
        existingId = byIdent.get(ident);
      }else if(!ident){
        const keyNG = (name+"||"+grade).toUpperCase();
        if(byNameGrade.has(keyNG)) existingId = byNameGrade.get(keyNG);
      }

      if(existingId){
        // ACTUALIZAR
        const st = db.students.find(s=>s.id===existingId);
        if(st){
          const oldIdent = (st.ident||"").trim();
          st.name = name;
          st.ident = ident;   // se respeta el archivo (si viene vacío, queda vacío)
          st.grade = grade;
          st.code = code;
          updated++;

          if(oldIdent && oldIdent !== ident) byIdent.delete(oldIdent);
          if(ident) byIdent.set(ident, st.id);

          const keyNG = (name+"||"+grade).toUpperCase();
          byNameGrade.set(keyNG, st.id);
        }
      }else{
        // CREAR
        const newId = uid("st");
        db.students.push({ id:newId, name, ident, grade, code });
        created++;

        if(ident) byIdent.set(ident, newId);
        const keyNG = (name+"||"+grade).toUpperCase();
        byNameGrade.set(keyNG, newId);
      }
    }

    saveDB();
    renderStudents();
    fillStudentSelect();
    renderGradebookRows();
    renderFinalArea();
    renderGenerator();

    alert(
      "Importación finalizada.\n\n" +
      `Creados: ${created}\n` +
      `Actualizados: ${updated}\n` +
      `Omitidos (sin nombre): ${skippedEmpty}`
    );

  }catch(err){
    console.error(err);
    alert("No se pudo importar el archivo. Verifica que sea XLSX válido y que tenga los encabezados requeridos.");
  }
}
/* =========================
   CRUD ÁREAS
   ========================= */
function addArea(){
  const name = (document.getElementById("arName").value||"").trim();
  const ihs = (document.getElementById("arIhs").value||"").trim();
  const teacher = (document.getElementById("arTeacher").value||"").trim();
  const obs = (document.getElementById("arObs").value||"").trim();
  if(!name){ alert("Ingresa el nombre del área."); return; }

  const area = { id: uid("ar"), name, ihs, teacher, obs };
  db.areas.push(area);
  if(!db.grades[area.id]) db.grades[area.id] = {};
  saveDB();
  clearAreaForm();
  renderAreas();
  fillAreaSelects();
  renderGradebook();
  renderFinalArea();
  renderGenerator();
}
function clearAreaForm(){
  document.getElementById("arName").value="";
  document.getElementById("arIhs").value="";
  document.getElementById("arTeacher").value="";
  document.getElementById("arObs").value="";
}
function updateArea(id){
  const a = db.areas.find(x=>x.id===id); if(!a) return;
  const name = prompt("Nombre del área:", a.name||""); if(name===null) return;
  const ihs = prompt("IHS:", a.ihs||""); if(ihs===null) return;
  const teacher = prompt("DOCENTE:", a.teacher||""); if(teacher===null) return;
  const obs = prompt("OBS:", a.obs||""); if(obs===null) return;
  a.name = (name||"").trim();
  a.ihs = (ihs||"").trim();
  a.teacher = (teacher||"").trim();
  a.obs = (obs||"").trim();
  saveDB();
  renderAreas(); fillAreaSelects(); renderGradebook(); renderFinalArea(); renderGenerator();
}
function deleteArea(id){
  const a = db.areas.find(x=>x.id===id); if(!a) return;
  if(!confirm(`¿Eliminar área?\n\n${a.name}\n\nSe eliminarán sus calificaciones asociadas.`)) return;
  db.areas = db.areas.filter(x=>x.id!==id);
  if(db.grades[id]) delete db.grades[id];
  saveDB();
  renderAreas(); fillAreaSelects(); renderGradebook(); renderFinalArea(); renderGenerator();
}
function renderAreas(){
  const tb = document.getElementById("areasTbody");
  const q = (document.getElementById("arFilter")?.value||"").toLowerCase().trim();
  let rows = db.areas.slice();
  if(q){
    rows = rows.filter(a =>
      (a.name||"").toLowerCase().includes(q) ||
      (a.ihs||"").toLowerCase().includes(q) ||
      (a.teacher||"").toLowerCase().includes(q) ||
      (a.obs||"").toLowerCase().includes(q)
    );
  }
  tb.innerHTML = rows.map(a=>`
    <tr>
      <td>${escapeHTML(a.name||"")}</td>
      <td class="center mono">${escapeHTML(a.ihs||"")}</td>
      <td>${escapeHTML(a.teacher||"")}</td>
      <td>${escapeHTML(a.obs||"")}</td>
      <td class="right">
        <button class="btn ghost" onclick="updateArea('${a.id}')">Editar</button>
        <button class="btn danger" onclick="deleteArea('${a.id}')">Eliminar</button>
      </td>
    </tr>
  `).join("");
}

/* =========================
   SELECTS
   ========================= */
function fillAreaSelects(){
  const selects = [document.getElementById("gbArea"), document.getElementById("faArea")].filter(Boolean);
  selects.forEach(sel=>{
    const current = sel.value;
    sel.innerHTML = db.areas.length
      ? db.areas.map(a=>`<option value="${a.id}">${escapeHTML(a.name||"")}</option>`).join("")
      : `<option value="">(No hay áreas)</option>`;
    if(current && db.areas.some(a=>a.id===current)) sel.value = current;
  });
  renderGBMeta();
}
function fillStudentSelect(){
  const sel = document.getElementById("genStudent");
  if(!sel) return;
  const current = sel.value;
  sel.innerHTML = db.students.length
    ? db.students.map(s=>`<option value="${s.id}">${escapeHTML(s.name||"")}${s.ident?(" · "+escapeHTML(s.ident)):""}</option>`).join("")
    : `<option value="">(No hay estudiantes)</option>`;
  if(current && db.students.some(s=>s.id===current)) sel.value = current;
}
function fillPeriodSelects(){
  const periods = db.settings.periods;
  const sel = document.getElementById("gbPeriod");
  if(!sel) return;
  const current = sel.value;
  sel.innerHTML = Array.from({length:periods}, (_,i)=>`<option value="${i+1}">PERIODO ${i+1}</option>`).join("");
  if(current) sel.value = current;
}
function fillPeriodSelectsGenerator(){
  const periods = db.settings.periods;
  const sel = document.getElementById("genPeriod");
  if(!sel) return;
  const current = sel.value;
  sel.innerHTML = Array.from({length:periods}, (_,i)=>`<option value="${i+1}">PERIODO ${i+1}</option>`).join("");
  if(current) sel.value = current;
  const mode = document.getElementById("genMode")?.value || "final";
  document.getElementById("genPeriodBox").classList.toggle("hide", mode!=="period");
}

/* =========================
   PLANILLA
   ========================= */
function ensureGrades(areaId, period){
  if(!db.grades[areaId]) db.grades[areaId] = {};
  if(!db.grades[areaId][period]) db.grades[areaId][period] = {};
}
function renderGBMeta(){
  const aId = document.getElementById("gbArea")?.value;
  const p = Number(document.getElementById("gbPeriod")?.value || 1);
  const a = db.areas.find(x=>x.id===aId);
  const cfg = getPeriodCfg(p);
  const txt = a ? `${a.name} · PERIODO ${p||"-"} · Estudiantes: ${db.students.length}` : "—";
  const el = document.getElementById("gbMeta");
  if(el) el.textContent = txt;

  const note = document.getElementById("gbCfgNote");
  if(note){
    const wp = Math.round(cfg.wPr*100), wc = Math.round(cfg.wCo*100), wa = Math.round(cfg.wAc*100);
    note.innerHTML = `
      <b>Configuración del PERIODO ${p}:</b>
      PROCEDIMENTAL (<b>${wp}%</b>) con <b>${cfg.prNotes}</b> notas (N1..N${cfg.prNotes}) + AE,
      CONCEPTUAL (<b>${wc}%</b>) con <b>${cfg.coNotes}</b> notas (N1..N${cfg.coNotes}) + AE,
      ACTITUDINAL (<b>${wa}%</b>) igual (DOC., EST. + ST) + REC.
      <br>ST = PROMEDIO:sobre las notas definidas. NOTA DEFINITIVA = ST(PR)+ ST(CO)+ ST(AC).
    `;
  }
}
function gradeCellInput(studentId, field, value){
  const aId = document.getElementById("gbArea").value;
  const period = Number(document.getElementById("gbPeriod").value);
  if(!aId || !period) return;

  ensureGrades(aId, period);
  if(!db.grades[aId][period][studentId]) db.grades[aId][period][studentId] = emptyRow();
  const row = db.grades[aId][period][studentId];

  if(field==="j" || field==="sj"){
    row[field] = toNum(value) ?? 0;
  }else{
    row[field] = value==="" ? null : toNum(value);
  }
  calcRowExcelLike(row, getPeriodCfg(period));
}
function renderGradebook(){
  fillAreaSelects();
  fillPeriodSelects();
  renderGBMeta();
  renderGradebookRows();
}
function renderGradebookRows(){
  const wrap = document.getElementById("gbTableWrap");
  if(!wrap) return;

  const aId = document.getElementById("gbArea").value;
  const period = Number(document.getElementById("gbPeriod").value);
  renderGBMeta();

  if(!aId){ wrap.innerHTML = `<div class="note">Crea un área primero (menú Áreas).</div>`; return; }
  if(!db.students.length){ wrap.innerHTML = `<div class="note">Agrega estudiantes primero (menú Estudiantes).</div>`; return; }

  ensureGrades(aId, period);
  const cfg = getPeriodCfg(period);

  const q = (document.getElementById("gbFilter")?.value||"").toLowerCase().trim();
  let students = db.students.slice();
  if(q){
    students = students.filter(s =>
      (s.name||"").toLowerCase().includes(q) ||
      (s.ident||"").toLowerCase().includes(q)
    );
  }

  const prCols = Array.from({length:cfg.prNotes}, (_,i)=> `<th class="center">NOTA${i+1}</th>`).join("");
  const coCols = Array.from({length:cfg.coNotes}, (_,i)=> `<th class="center">NOTA${i+1}</th>`).join("");

  const head = `
    <div class="twrap">
    <table>
      <thead>
        <tr>
          <th rowspan="2">Nº</th>
          <th rowspan="2">APELLIDOS Y NOMBRES</th>
          <th colspan="3" class="center">INASISTENCIA</th>

          <th colspan="${cfg.prNotes + 3}" class="center">PROCEDIMENTAL (${Math.round(cfg.wPr*100)}%)</th>
          <th colspan="${cfg.coNotes + 3}" class="center">CONCEPTUAL (${Math.round(cfg.wCo*100)}%)</th>
          <th colspan="5" class="center">ACTITUDINAL (${Math.round(cfg.wAc*100)}%)</th>

          <th rowspan="2" class="center">NOTA<br>DEFINITIVA</th>
          <th rowspan="2" class="center">ESCALA<br>NACIONAL</th>
        </tr>
        <tr>
          <th class="center">J</th>
          <th class="center">SJ</th>
          <th class="center">TOTAL</th>

          ${prCols}<th class="center">AUTO-EVAL</th><th class="center">SUMA-TOTAL</th><th class="center">RECUPERACIÓN</th>
          ${coCols}<th class="center">AUTO-EVAL</th><th class="center">SUMA-TOTAL</th><th class="center">RECUPERACIÓN</th>

          <th class="center">DOC.</th><th class="center">EST.</th><th class="center">ST</th><th class="center">0.${Math.round(cfg.wAc*100)/10}</th><th class="center">REC.</th>
        </tr>
      </thead>
      <tbody>
  `;

  const body = students.map((s,i)=>{
    const row = db.grades[aId][period][s.id] || emptyRow();
    calcRowExcelLike(row, cfg);

    const totalIna = (toNum(row.j)||0) + (toNum(row.sj)||0);
    const nota = row._computed?.nota ?? null;

    const prInputs = Array.from({length:cfg.prNotes}, (_,k)=>{
      const idx = k+1;
      const f = noteField("pr", idx);
      return numInput(s.id, f, row[f]);
    }).join("");

    const coInputs = Array.from({length:cfg.coNotes}, (_,k)=>{
      const idx = k+1;
      const f = noteField("co", idx);
      return numInput(s.id, f, row[f]);
    }).join("");

    return `
      <tr data-st="${s.id}">
        <td class="mono">${i+1}</td>
        <td>${escapeHTML(s.name||"")}</td>

        <td><input class="num" value="${escapeHTML(String(row.j??0))}" oninput="gradeCellInput('${s.id}','j',this.value); updateRowView('${s.id}')"></td>
        <td><input class="num" value="${escapeHTML(String(row.sj??0))}" oninput="gradeCellInput('${s.id}','sj',this.value); updateRowView('${s.id}')"></td>
        <td class="center mono" id="ina_${s.id}">${totalIna}</td>

        ${prInputs}
        ${numInput(s.id,'pr_ae',row.pr_ae)}
        ${numInput(s.id,'pr_st',row._computed?.pr_st)}
        ${numInput(s.id,'pr_rec',row.pr_rec)}

        ${coInputs}
        ${numInput(s.id,'co_ae',row.co_ae)}
        ${numInput(s.id,'co_st',row._computed?.co_st)}
        ${numInput(s.id,'co_rec',row.co_rec)}

        ${numInput(s.id,'ac_doc',row.ac_doc)}
        ${numInput(s.id,'ac_est',row.ac_est)}
        ${numInput(s.id,'ac_st',row._computed?.ac_st)}
        <td class="center mono" id="acw_${s.id}">${isNum(row._computed?.ac_w)?fmt(row._computed.ac_w,2):""}</td>
        ${numInput(s.id,'ac_rec',row.ac_rec)}

        <td class="center mono" id="nota_${s.id}">${isNum(nota)?fmt(nota,2):""}</td>
        <td class="center" id="esc_${s.id}">${isNum(nota)?escapeHTML(escalaNacional_D(nota)):" "}</td>
      </tr>
    `;
  }).join("");

  wrap.innerHTML = head + body + `</tbody></table></div>`;
}

/* =========================
   IMPORTAR PLANILLAS (XLSX) — Solo vista "Planillas por Área"
   Formato esperado: plantilla tipo AREAS.xlsx (Nº, Apellidos y Nombres, Inasistencia, Procedimental, Conceptual, Actitudinal)
   ========================= */
function openPlanillasImport(){
  const aId = document.getElementById("gbArea")?.value;
  const period = Number(document.getElementById("gbPeriod")?.value);
  if(!aId){ toast("Selecciona un área antes de importar."); return; }
  if(!period){ toast("Selecciona el periodo antes de importar."); return; }
  const el = document.getElementById("planillasFile");
  if(!el){ toast("No se encontró el selector de archivo."); return; }
  el.value = "";
  el.click();
}

function normName(s){
  return (String(s||""))
    .normalize("NFD").replace(/[̀-ͯ]/g,"")
    .replace(/\s+/g," ")
    .trim().toUpperCase();
}

function pickSheetForPeriod(wb, period){
  const p = Number(period)||1;
  const candidates = [
    `P${p}`, `P ${p}`, `PERIODO ${p}`, `PERÍODO ${p}`, `PERIODO_${p}`, `PERÍODO_${p}`,
    `PERIODO${p}`, `PERÍODO${p}`, `Periodo ${p}`, `P${p}`.toLowerCase()
  ].map(x=>String(x));
  const byExact = wb.SheetNames.find(n=>candidates.includes(String(n).trim()));
  if(byExact) return byExact;

  // heurística: que contenga "P1", "P2"...
  const re = new RegExp(`(^|\b)P\s*${p}(\b|$)`, "i");
  const byLike = wb.SheetNames.find(n=>re.test(String(n)));
  return byLike || wb.SheetNames[0];
}

function showPlanillasImportMsg(html, isError=false){
  const el = document.getElementById("planillasImportMsg");
  if(!el) return;
  el.style.display = "block";
  el.style.borderColor = isError ? "rgba(255,107,107,.55)" : "rgba(90,169,255,.35)";
  el.innerHTML = html;
}

async function handlePlanillasImport(file){
  if(!file) return;
  try{
    const aId = document.getElementById("gbArea")?.value;
    const period = Number(document.getElementById("gbPeriod")?.value);
    if(!aId || !period) return;

    if(!window.XLSX){
      showPlanillasImportMsg("No se encontró la librería XLSX (SheetJS).", true);
      return;
    }

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:"array"});
    const sheetName = pickSheetForPeriod(wb, period);
    const ws = wb.Sheets[sheetName];
    if(!ws){
      showPlanillasImportMsg(`No se encontró la hoja "${escapeHTML(sheetName)}" en el archivo.`, true);
      return;
    }

    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, blankrows:false});
    if(!rows || !rows.length){
      showPlanillasImportMsg("La planilla no tiene datos.", true);
      return;
    }

    // localizar encabezado (fila con "Nº" y "APELLIDOS Y NOMBRES")
    const headerRow = rows.findIndex(r=>{
      const a = (r?.[0] ?? "");
      const b = (r?.[1] ?? "");
      const A = String(a).trim().toUpperCase();
      const B = String(b).trim().toUpperCase();
      return (A==="Nº" || A==="N°" || A==="NO" || A==="N" || A==="Nº.") && B.includes("APELLIDOS");
    });

    const startIdx0 = headerRow >= 0 ? headerRow + 2 : 0; // +2: subencabezado
    let i0 = startIdx0;
    while(i0 < rows.length){
      const name = rows[i0]?.[1];
      const any = (rows[i0]||[]).some(v=>v!==null && v!=="" && v!==undefined);
      if(any && String(name||"").trim()) break;
      i0++;
    }

    // Detección opcional de columnas extra (Identificación/Código) en la planilla.
    const header0 = headerRow >= 0 ? (rows[headerRow] || []) : (rows[0] || []);
    const header1 = headerRow >= 0 ? (rows[headerRow+1] || []) : [];
    const u = v => String(v??"").trim().toUpperCase();
    const findHeaderIndex = (kws)=>{
      const arrs = [header0, header1];
      for(const arr of arrs){
        for(let i=0;i<(arr||[]).length;i++){
          const cell = u(arr[i]);
          if(!cell) continue;
          if(kws.some(k => cell.includes(k))) return i;
        }
      }
      return -1;
    };
    const identCol = findHeaderIndex(["IDENT", "DOCUMENTO", "DOC", "CC", "TI"]);
    const codeCol  = findHeaderIndex(["CÓDIGO", "CODIGO", "COD."]);
    const baseShift = (identCol === 2 || codeCol === 2) ? 1 : 0;

    const looksLikeIdent = (v)=>{
      const d = String(v??"").replace(/\D/g,"");
      return d.length >= 6;
    };
    const looksLikeCode = (v)=>{
      const s = String(v??"").trim();
      return s.length >= 3 && /[A-Za-z]/.test(s);
    };

    // índices de estudiantes (prioridad: identificación/código; fallback: nombre)
    const nameToIds = new Map();
    const identToIds = new Map();
    const codeToIds = new Map();

    const normIdent = (v)=> String(v||"").replace(/\D/g,"").trim();
    const normCode  = (v)=> String(v||"").trim().toUpperCase().replace(/\s+/g,"");

    (db.students||[]).forEach(s=>{
      const nk = normName(s.name);
      if(nk){
        if(!nameToIds.has(nk)) nameToIds.set(nk, []);
        nameToIds.get(nk).push(s.id);
      }
      const ik = normIdent(s.ident);
      if(ik){
        if(!identToIds.has(ik)) identToIds.set(ik, []);
        identToIds.get(ik).push(s.id);
      }
      const ck = normCode(s.code);
      if(ck){
        if(!codeToIds.has(ck)) codeToIds.set(ck, []);
        codeToIds.get(ck).push(s.id);
      }
    });

ensureGrades(aId, period);
    const cfg = getPeriodCfg(period);

    let imported = 0, updatedFields = 0;
    const unmatched = [];
    const duplicates = [];
    const touched = [];

    for(let r=i0; r<rows.length; r++){
      const row = rows[r] || [];
      const name = row[1];
      const key = normName(name);

      // Si la planilla trae una columna extra (Identificación/Código) después del nombre,
      // desplazamos los índices para leer correctamente las notas.
      let shift = baseShift;
      let rowIdent = (identCol >= 0 ? row[identCol] : "");
      let rowCode  = (codeCol  >= 0 ? row[codeCol]  : "");

      // Heurística: si no detectamos columnas por encabezado y en la col 2 hay algo tipo ID/Código, asumimos shift=1.
      if(shift === 0 && !rowIdent && !rowCode){
        const c2 = row[2];
        const c3 = row[3];
        if((looksLikeIdent(c2) || looksLikeCode(c2)) && !looksLikeIdent(c3)){
          shift = 1;
          if(looksLikeIdent(c2)) rowIdent = c2;
          else rowCode = c2;
        }
      }

      const baseCols = [2,3,5,6,7,8,9,11,12,13,14,15,17,18,19,22];
      const hasSomething = baseCols.some(ix=>{
        const v = row[ix + shift];
        return v!==null && v!=="" && v!==undefined;
      });

      if(!key && !hasSomething) continue;
      if(!key){ continue; }

      // Resolver estudiante (prioridad: Identificación, luego Código, luego Nombre)
      const identKey = rowIdent ? normIdent(rowIdent) : "";
      const codeKey  = rowCode  ? normCode(rowCode)  : "";

      let ids = null;
      let matchMode = "nombre";
      if(identKey && identToIds.has(identKey)){
        ids = identToIds.get(identKey);
        matchMode = "identificación";
      }else if(codeKey && codeToIds.has(codeKey)){
        ids = codeToIds.get(codeKey);
        matchMode = "código";
      }else{
        ids = nameToIds.get(key);
      }

      if(!ids || !ids.length){
        const tag = identKey ? ` (ID: ${identKey})` : (codeKey ? ` (COD: ${codeKey})` : "");
        unmatched.push(String(name||"").trim() + tag);
        continue;
      }
      if(ids.length>1){
        const tag = identKey ? ` [ID duplicada: ${identKey}]` : (codeKey ? ` [COD duplicado: ${codeKey}]` : "");
        duplicates.push(String(name||"").trim() + tag + ` (por ${matchMode})`);
      }

      const studentId = ids[0];

      let gr = db.grades[aId][period][studentId] || emptyRow();

      // helpers: solo sobreescribe si viene dato
      const setIf = (field, value)=>{
        const v = toNum(value);
        if(value===null || value==="" || value===undefined) return;
        if(field==="j" || field==="sj"){
          gr[field] = Math.max(0, Math.round(toNum(value) ?? 0));
        }else{
          gr[field] = isNum(v) ? clamp(v,0,5) : gr[field];
        }
        updatedFields++;
      };

      setIf("j",  row[2+shift]);  // INASISTENCIA J
      setIf("sj", row[3+shift]);  // INASISTENCIA SJ

      // PROCEDIMENTAL
      setIf("pr_n1", row[5+shift]);
      setIf("pr_n2", row[6+shift]);
      setIf("pr_n3", row[7+shift]);
      setIf("pr_n4", row[8+shift]);
      setIf("pr_ae", row[9+shift]);
      // row[10] = ST (opcional), si lo traen, se respeta:
      if(row[10+shift]!==null && row[10+shift]!=="" && row[10+shift]!==undefined) setIf("pr_st", row[10+shift]);
      setIf("pr_rec", row[11+shift]);

      // CONCEPTUAL
      setIf("co_n1", row[12+shift]);
      setIf("co_n2", row[13+shift]);
      setIf("co_n3", row[14+shift]);
      setIf("co_ae", row[15+shift]);
      if(row[16+shift]!==null && row[16+shift]!=="" && row[16+shift]!==undefined) setIf("co_st", row[16+shift]);
      setIf("co_rec", row[17+shift]);

      // ACTITUDINAL
      setIf("ac_doc", row[18+shift]);
      setIf("ac_est", row[19+shift]);
      if(row[20+shift]!==null && row[20+shift]!=="" && row[20+shift]!==undefined) setIf("ac_st", row[20+shift]);
      setIf("ac_rec", row[22+shift]);

      gr = calcRowExcelLike(gr, cfg);
      db.grades[aId][period][studentId] = gr;
      imported++;
      touched.push(studentId);
    }

    saveDB();
    renderGradebookRows();
    renderGBMeta();

    const msg = `
      <b>Importación completada</b><br/>
      Archivo: <span class="mono">${escapeHTML(file.name)}</span><br/>
      Hoja: <span class="mono">${escapeHTML(sheetName)}</span><br/>
      Registros importados: <b>${imported}</b> · Celdas actualizadas: <b>${updatedFields}</b>
      ${duplicates.length ? `<div style="margin-top:8px"><b>⚠ Nombres duplicados en tu lista de estudiantes:</b> ${escapeHTML(duplicates.slice(0,8).join(", "))}${duplicates.length>8?"…":""}</div>` : ""}
      ${unmatched.length ? `<div style="margin-top:8px"><b>No encontrados (por nombre):</b> ${escapeHTML(unmatched.slice(0,12).join(", "))}${unmatched.length>12?"…":""}</div>` : ""}
      <div style="margin-top:8px;color:var(--muted)">Sugerencia: para mejores coincidencias, asegúrate de que los nombres en la planilla coincidan exactamente con “Estudiantes”.</div>
    `;
    showPlanillasImportMsg(msg, false);
    toast("Planillas importadas.");
  }catch(err){
    console.error(err);
    showPlanillasImportMsg(`Error importando planillas: ${escapeHTML(err?.message || String(err))}`, true);
    toast("Error al importar planillas.");
  }
}

function numInput(studentId, field, val){
  const v = (val===null || val===undefined) ? "" : (isNum(toNum(val)) ? String(toNum(val)) : (isNum(val)?String(val):""));
  return `<td><input class="num" value="${escapeHTML(v)}" oninput="gradeCellInput('${studentId}','${field}',this.value); updateRowView('${studentId}')"></td>`;
}
function updateRowView(studentId){
  const aId = document.getElementById("gbArea").value;
  const period = Number(document.getElementById("gbPeriod").value);
  if(!aId || !period) return;

  ensureGrades(aId, period);
  const cfg = getPeriodCfg(period);
  const row = db.grades[aId][period][studentId] || emptyRow();
  calcRowExcelLike(row, cfg);

  const ina = (toNum(row.j)||0) + (toNum(row.sj)||0);
  const nota = row._computed?.nota ?? null;

  const inaEl = document.getElementById("ina_"+studentId); if(inaEl) inaEl.textContent = ina;
  const notaEl = document.getElementById("nota_"+studentId); if(notaEl) notaEl.textContent = isNum(nota) ? fmt(nota,2) : "";
  const escEl = document.getElementById("esc_"+studentId); if(escEl) escEl.textContent = isNum(nota) ? escalaNacional_D(nota) : " ";

  const acw = document.getElementById("acw_"+studentId); if(acw) acw.textContent = isNum(row._computed?.ac_w)?fmt(row._computed.ac_w,2):"";
}
function recalcAllVisible(){
  document.querySelectorAll("#gbTableWrap tbody tr").forEach(tr=> updateRowView(tr.dataset.st));
}
function saveAllVisible(){
  saveDB();
  alert("Guardado. (Persistido en este navegador)");
}

/* =========================
   FINAL_AREA.xlsx
   ========================= */
function computeAreaFinalForStudent(areaId, studentId){
  const periods = db.settings.periods;
  let inaJ=0, inaSJ=0;

  const per = [];
  for(let p=1;p<=periods;p++){
    const row = db.grades?.[areaId]?.[p]?.[studentId];
    if(row){
      calcRowExcelLike(row, getPeriodCfg(p));
      inaJ += (toNum(row.j)||0);
      inaSJ += (toNum(row.sj)||0);
      const nota = row._computed?.nota ?? null;
      const hasRec = [row.pr_rec, row.co_rec, row.ac_rec].map(toNum).some(v => isNum(v) && v > 0.9);
      per.push({p, nota, mej: hasRec, des: isNum(nota)? escalaNacional_D(nota) : " "});
    }else{
      per.push({p, nota:null, des:" "});
    }
  }

  const inaTotal = inaJ + inaSJ;
  const notaDef = avgIfGT09_periodNotes(per.map(x=>x.nota));
  const desDef = isNum(notaDef) ? escalaNacional_D(notaDef) : " ";

  return { inaJ, inaSJ, inaTotal, per, notaDef, desDef };
}

function renderFinalArea(){
  fillAreaSelects();
  const wrap = document.getElementById("faTableWrap");
  if(!wrap) return;

  const areaId = document.getElementById("faArea").value;
  const q = (document.getElementById("faFilter")?.value||"").toLowerCase().trim();

  if(!areaId){ wrap.innerHTML = `<div class="note">Crea un área primero (menú Áreas).</div>`; return; }
  if(!db.students.length){ wrap.innerHTML = `<div class="note">Agrega estudiantes primero (menú Estudiantes).</div>`; return; }

  let students = db.students.slice();
  if(q){
    students = students.filter(s =>
      (s.name||"").toLowerCase().includes(q) ||
      (s.ident||"").toLowerCase().includes(q)
    );
  }

  const periods = db.settings.periods;
  const thPeriods = Array.from({length:periods}, (_,i)=>`
    <th class="center">P${i+1}<br>NOTA</th>
    <th class="center">P${i+1}<br>DESEMPEÑO</th>
  `).join("");

  const head = `
    <div class="twrap">
    <table>
      <thead>
        <tr>
          <th>Nº</th>
          <th>APELLIDOS Y NOMBRES</th>
          <th class="center">J</th>
          <th class="center">SJ</th>
          <th class="center">TOTAL</th>
          ${thPeriods}
          <th class="center">NOTA<br>DEFINITIVA</th>
          <th class="center">ESCALA<br>NACIONAL</th>
        </tr>
      </thead>
      <tbody>
  `;

  const body = students.map((s,i)=>{
    const fin = computeAreaFinalForStudent(areaId, s.id);
    const perCells = fin.per.map(x=>`
      <td class="center mono">${isNum(x.nota)?fmt(x.nota,2):""}</td>
      <td class="center">${escapeHTML(x.des||" ")}</td>
    `).join("");
    return `
      <tr>
        <td class="mono">${i+1}</td>
        <td>${escapeHTML(s.name||"")}</td>
        <td class="center mono">${fin.inaJ}</td>
        <td class="center mono">${fin.inaSJ}</td>
        <td class="center mono">${fin.inaTotal}</td>
        ${perCells}
        <td class="center mono">${isNum(fin.notaDef)?fmt(fin.notaDef,2):""}</td>
        <td class="center">${escapeHTML(fin.desDef||" ")}</td>
      </tr>
    `;
  }).join("");

  wrap.innerHTML = head + body + `</tbody></table></div>`;
}

/* =========================
   GENERADOR (BOLETÍN)
   ========================= */
function renderGenerator(){
  fillStudentSelect();
  fillPeriodSelectsGenerator();

  const wrap = document.getElementById("genTableWrap");
  if(!wrap) return;

  const studentId = document.getElementById("genStudent").value;
  const mode = document.getElementById("genMode").value;
  fillPeriodSelectsGenerator();

  if(!studentId){
    wrap.innerHTML = `<div class="note">Agrega y selecciona un estudiante.</div>`;
    return;
  }
  const st = db.students.find(s=>s.id===studentId);
  const periods = db.settings.periods;
  const grade = (st.grade || db.settings.gradeGlobal || "");
  const grupo = (db.settings.grupo||"");
  const year = (db.settings.year||"");
  const jornada = (db.settings.jornada||"");
  const fecha = (db.settings.fechaBoletin||"");

  
// Promovido / Al grado por estudiante (no global)
if(!st.boletin) st.boletin = {};
const promBo = (year ? getBoletinPromovidoForStudent(st) : "");
const alGrBo = (year ? getBoletinAlGradoForStudent(st) : "");

// Sincroniza la UI de Promovido/Al grado cuando cambias de estudiante
const proEl = document.getElementById("setPromovido");
if(proEl) proEl.value = (year ? promBo : "");
const alEl = document.getElementById("setAlGrado");
if(alEl) alEl.value = alGrBo || "";
/* =========================
     POSICIÓN DE GRUPO (promedio total por periodo)
     - Calcula el promedio del estudiante en TODAS las áreas en un periodo (AVERAGEIF >0,9).
     - Posición 1 = mayor promedio del grupo (mismo GRADO).
     ========================= */
  const gradeKey = String(grade||"").trim();

  function _sameGrade(s){
    if(!gradeKey) return true;
    const g = String(s.grade || db.settings.gradeGlobal || "").trim();
    return g === gradeKey;
  }

  function _avgAllAreasForPeriod(stuId, p){
    const notas = db.areas.map(a=>{
      const row = db.grades?.[a.id]?.[p]?.[stuId];
      if(row){
        calcRowExcelLike(row, getPeriodCfg(p)); // asegura _computed
        const nota = row._computed?.nota ?? null;
        return (isNum(nota) && nota > 0.9) ? nota : null;
      }
      return null;
    });
    return avgIfGT09_periodNotes(notas);
  }

  function _posGrupoForPeriod(stuId, p){
    const peers = db.students.filter(_sameGrade);
    const list = peers.map(s=>{
      const avg = _avgAllAreasForPeriod(s.id, p);
      return { id:s.id, avg };
    }).filter(x=>isNum(x.avg));

    if(!list.some(x=>x.id===stuId)) return "";

    list.sort((a,b)=> b.avg - a.avg);

    let rank = 0;
    let prev = null;
    for(let i=0;i<list.length;i++){
      const v = list[i];
      if(prev === null || Math.abs(v.avg - prev) > 1e-9){
        rank = i + 1;
        prev = v.avg;
      }
      if(v.id === stuId) return String(rank);
    }
    return "";
  }

  function _posGrupoResumen(){
    const parts = [];
    for(let p=1;p<=periods;p++){
      const r = _posGrupoForPeriod(studentId, p);
      parts.push(`P${p}: ${r || "—"}`);
    }
    return parts.join(" · ");
  }


  const ieName = (db.settings.ieName || db.settings.institution || "INSTITUCIÓN EDUCATIVA").trim();
  const nit = (db.settings.nit||"").trim();
  const dane = (db.settings.dane||"").trim();
  const place = (db.settings.place||"").trim();
  const instLine = (db.settings.institution||"").trim();

  const shield = db.settings.images?.shield || "";

  const rectorName = (db.settings.rectorName||"").trim();
  const dirName = (db.settings.dirName||"").trim();
  const rectorSign = db.settings.images?.rectorSign || "";
  const rectorSeal = db.settings.images?.rectorSeal || "";
  const dirSign = db.settings.images?.dirSign || "";
  const dirSeal = db.settings.images?.dirSeal || "";

  const items = db.areas.map(a=>{
    const fin = computeAreaFinalForStudent(a.id, studentId);
    const anual = avgSimple(fin.per.map(x=>x.nota));
    const desAnual = isNum(anual) ? escalaNacional_SIN_D(anual) : " ";
    const perOut = fin.per.map(x=>{
      const des = isNum(x.nota) ? escalaNacional_SIN_D(x.nota) : " ";
      return { nota:x.nota, mej: !!x.mej, des };
    });
    const ihsNum = toNum(a.ihs);
    return { area:a, fin, anual, desAnual, perOut, ihsNum:isNum(ihsNum)?ihsNum:0 };
  });

  const totalIHS = items.reduce((acc,x)=> acc + (x.ihsNum||0), 0);
  const promPeriodos = Array.from({length:periods}, (_,i)=>{
    const notas = items.map(it => it.perOut[i]?.nota).filter(isNum);
    return avgSimple(notas);
  });
  const promAnual = avgSimple(items.map(it=>it.anual).filter(isNum));

  const sumJ = items.reduce((acc,it)=> acc + (it.fin.inaJ||0), 0);
  const sumSJ = items.reduce((acc,it)=> acc + (it.fin.inaSJ||0), 0);
  const sumTotal = sumJ + sumSJ;

  const desGlobal = isNum(promAnual) ? escalaNacional_SIN_D(promAnual) : " ";

  const periodHeaders = Array.from({length:periods}, (_,i)=>{
    const lbl = (i===0) ? "PRIMER PERIODO" : (i===1) ? "SEGUNDO PERIODO" : (i===2) ? "TERCER PERIODO" : `CUARTO PERIODO`;
    return `<th colspan="3">${lbl}</th>`;
  }).join("");

  const periodSubHeaders = Array.from({length:periods}, ()=>`
    <th>MEJ.</th><th>NOTA</th><th>DESEMPEÑO</th>
  `).join("");

  let rowsHTML = "";
  items.forEach(it=>{
    rowsHTML += `
      <tr>
        <td>${escapeHTML(it.area.name||"")}</td>
        <td class="b-center">${escapeHTML(String(it.area.ihs||""))}</td>
        <td>${escapeHTML(it.area.teacher||"")}</td>

        <td class="b-center">${it.fin.inaJ||0}</td>
        <td class="b-center">${it.fin.inaSJ||0}</td>
        <td class="b-center">${it.fin.inaTotal||0}</td>

        ${it.perOut.map(p=>{
          const mej = p.mej ? "X" : "";
          return `
            <td class="b-center">${escapeHTML(mej)}</td>
            <td class="b-center">${isNum(p.nota)?fmt(p.nota,2):""}</td>
            <td class="b-center">${escapeHTML(p.des||" ")}</td>
          `;
        }).join("")}

        <td class="b-center">${isNum(it.anual)?fmt(it.anual,2):""}</td>
        <td class="b-center">${escapeHTML(it.desAnual||" ")}</td>
      </tr>
    `;
  });

  function boletinHeaderHTML(extraPeriodoText){
    return `
      <div class="b-head">
        <div class="b-escudo">
          ${shield ? `<img src="${shield}" alt="Escudo">` : `<span style="font-size:10px;color:#000">ESCUDO</span>`}
        </div>
        <div class="b-headtext">
          <div class="b-inst">${escapeHTML(ieName)}</div>
          <div class="b-sub">${escapeHTML(instLine)}</div>
          <div class="b-sub">${escapeHTML(place)}</div>
          <div class="b-subline">
            ${nit ? `<b>NIT:</b> ${escapeHTML(nit)} &nbsp;&nbsp;` : ``}
            ${dane ? `<b>DANE:</b> ${escapeHTML(dane)}` : ``}
          </div>
          <div class="b-title">INFORME DE VALORACIÓN DEL DESEMPEÑO</div>
          ${extraPeriodoText ? `<div class="b-subline">${extraPeriodoText}</div>` : ``}
        </div>
      </div>
    `;
  }

  function firmasHTML(){
    return `
      <div class="b-signgrid">
        <div class="b-signbox">
          <div class="b-sigtop">
            ${(rectorSign||rectorSeal) ? `
              ${rectorSign ? `<img src="${rectorSign}" alt="Firma Rector">` : ``}
              ${rectorSeal ? `<img src="${rectorSeal}" alt="Sello Rector">` : ``}
            ` : ``}
          </div>
          <div class="b-sigline"></div>
          <div class="b-name">${escapeHTML(rectorName || "")}</div>
          <div class="b-role b-muted">RECTOR(A)</div>
        </div>

        <div class="b-signbox">
          <div class="b-sigtop">
            ${(dirSign||dirSeal) ? `
              ${dirSign ? `<img src="${dirSign}" alt="Firma Director de Grupo">` : ``}
              ${dirSeal ? `<img src="${dirSeal}" alt="Sello Director de Grupo">` : ``}
            ` : ``}
          </div>
          <div class="b-sigline"></div>
          <div class="b-name">${escapeHTML(dirName || "")}</div>
          <div class="b-role b-muted">DOCENTE DIRECTOR(A) DE GRUPO</div>
        </div>
      </div>
    `;
  }

  if(mode==="period"){
    const p = Number(document.getElementById("genPeriod").value);
    const idx = p-1;

    const posGrupo = _posGrupoForPeriod(studentId, p);

    const pd = (db.settings.periodDates && db.settings.periodDates[p]) ? db.settings.periodDates[p] : {};
    const sD = fmtISOToDMY(pd.start);
    const eD = fmtISOToDMY(pd.end);
    const fechas = (sD || eD) ? ` &nbsp;&nbsp; <b>FECHAS:</b> ${escapeHTML(sD)}${(sD&&eD)?' - ':''}${escapeHTML(eD)}` : "";

    const head = `
      <div class="boletin">
        ${boletinHeaderHTML(`<b>PERIODO:</b> ${escapeHTML(String(p))}${fechas}`)}

        <div class="b-grid">
          <div class="b-box">
            <div class="b-row"><b>ESTUDIANTE:</b><span>${escapeHTML(st.name||"")}</span></div>
            <div class="b-row"><b>IDENTIFICACION:</b><span>${escapeHTML(st.ident||"")}</span></div>
            <div class="b-row"><b>GRADO:</b><span>${escapeHTML(grade||"")}</span></div>
          </div>
          <div class="b-box">
            <div class="b-row"><b>AÑO LECTIVO:</b><span>${escapeHTML(year)}</span></div>
            <div class="b-row"><b>JORNADA:</b><span>${escapeHTML(jornada)}</span></div>
            <div class="b-row"><b>GRUPO:</b><span>${escapeHTML(grupo)}</span></div>
          </div>
          <div class="b-box">
            <div class="b-row"><b>FECHA:</b><span>${escapeHTML(fecha)}</span></div>
            <div class="b-row"><b>PERIODO:</b><span>${escapeHTML(String(p))}</span></div>
            <div class="b-row"><b>POSICIÓN GRUPO:</b><span>${escapeHTML(posGrupo)}</span></div>
          </div>
        </div>

        <table class="b-table">
          <thead>
            <tr>
              <th>AREA Y/O ASIGNATURA</th>
              <th>IHS</th>
              <th>DOCENTE</th>
              <th colspan="3">INASISTENCIA</th>
              <th>MEJ.</th>
              <th>NOTA</th>
              <th>DESEMPEÑO</th>
            </tr>
            <tr>
              <th></th><th></th><th></th>
              <th>J</th><th>SJ</th><th>TOTAL</th>
              <th></th><th></th><th></th>
            </tr>
          </thead>
          <tbody>
            ${items.map(it=>{
              const po = it.perOut[idx] || {nota:null, des:" "};
              return `
                <tr>
                  <td>${escapeHTML(it.area.name||"")}</td>
                  <td class="b-center">${escapeHTML(String(it.area.ihs||""))}</td>
                  <td>${escapeHTML(it.area.teacher||"")}</td>
                  <td class="b-center">${it.fin.inaJ||0}</td>
                  <td class="b-center">${it.fin.inaSJ||0}</td>
                  <td class="b-center">${it.fin.inaTotal||0}</td>
                  <td class="b-center"></td>
                  <td class="b-center">${isNum(po.nota)?fmt(po.nota,2):""}</td>
                  <td class="b-center">${escapeHTML(po.des||" ")}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>

        <div class="b-footgrid">
          <div class="b-bigline">
            <div class="b-row"><b>TIPO DE FALTAS</b><span class="b-muted">I 0 &nbsp;&nbsp; II 0 &nbsp;&nbsp; III 0</span></div>
            <div class="b-row"><b>TOTAL IHS</b><span class="b-muted">${totalIHS}</span></div>
            <div class="b-row"><b>INASISTENCIA (J / SJ / TOTAL)</b><span class="b-muted">${sumJ} / ${sumSJ} / ${sumTotal}</span></div>
            <div class="b-row"><b>PROMEDIO PERIODO</b><span class="b-muted">${isNum(promPeriodos[idx])?fmt(promPeriodos[idx],2):"0,00"} &nbsp;&nbsp; ${escapeHTML(isNum(promPeriodos[idx])?escalaNacional_SIN_D(promPeriodos[idx]):"BAJO")}</span></div>
          </div>

          <div class="b-legend">
            <h3>Descripciones generales de desempeño (Decreto 1290 de 2009)</h3>
            <div class="lrow"><b>1,00 - 2,99</b><span>DESEMPEÑO BAJO: No supera los desempeños previstos en el área.</span></div>
            <div class="lrow"><b>3,00 - 3,99</b><span>DESEMPEÑO BÁSICO: Logra lo mínimo; debe mejorar para alcanzar niveles esperados.</span></div>
            <div class="lrow"><b>4,00 - 4,59</b><span>DESEMPEÑO ALTO: Alcanza la mayoría de desempeños; buen nivel de desarrollo.</span></div>
            <div class="lrow"><b>4,60 - 5,00</b><span>DESEMPEÑO SUPERIOR: Alcanza la totalidad; cumple de manera cabal.</span></div>
            <div class="b-note b-muted">X: presentó mejoramiento · X: debió presentarlo y no lo hizo</div>
          </div>
        </div>

        <div class="b-footgrid">
          <div class="b-bigline">
            <div class="b-row"><b>PROMOVIDO:</b><span>${escapeHTML((db.settings.year ? (db.settings.promovido||"SI") : ""))} &nbsp;&nbsp; <b> </b> &nbsp;&nbsp; <b>AL GRADO:</b> ${escapeHTML(db.settings.alGrado||"")}</span></div>
            <div class="b-note"><b>ACTIVIDADES DE MEJORAMIENTO:</b><div class="b-muted" style="margin-top:6px;white-space:pre-wrap">${escapeHTML(db.settings.actMej||"")}</div></div>
          </div>
          <div class="b-bigline">
            <div class="b-note"><b>OBSERVACIONES:</b><div class="b-muted" style="margin-top:6px;white-space:pre-wrap">${escapeHTML(db.settings.obsBoletin||"")}</div></div>
          </div>
        </div>

        ${firmasHTML()}
      </div>
    `;
    wrap.innerHTML = head;
    return;
  }

  const posGrupo = _posGrupoResumen();

  const html = `
    <div class="boletin">
      ${boletinHeaderHTML("")}

      <div class="b-grid">
        <div class="b-box">
          <div class="b-row"><b>ESTUDIANTE:</b><span>${escapeHTML(st.name||"")}</span></div>
          <div class="b-row"><b>IDENTIFICACION:</b><span>${escapeHTML(st.ident||"")}</span></div>
          <div class="b-row"><b>GRADO:</b><span>${escapeHTML(grade||"")}</span></div>
        </div>
        <div class="b-box">
          <div class="b-row"><b>AÑO LECTIVO:</b><span>${escapeHTML(year)}</span></div>
          <div class="b-row"><b>JORNADA:</b><span>${escapeHTML(jornada)}</span></div>
          <div class="b-row"><b>GRUPO:</b><span>${escapeHTML(grupo)}</span></div>
        </div>
        <div class="b-box">
          <div class="b-row"><b>FECHA:</b><span>${escapeHTML(fecha)}</span></div>
          <div class="b-row"><b>PERIODOS:</b><span>${periods}</span></div>
          <div class="b-row"><b>POSICIÓN GRUPO:</b><span>${escapeHTML(posGrupo)}</span></div>
        </div>
      </div>

      <table class="b-table">
        <thead>
          <tr>
            <th rowspan="2">AREA Y/O ASIGNATURA</th>
            <th rowspan="2">IHS</th>
            <th rowspan="2">DOCENTE</th>
            <th colspan="3">INASISTENCIA</th>
            ${periodHeaders}
	    <th rowspan="2" style="text-align:center; vertical-align:middle;">
   	               DEFINITIVA<br>ANUAL
                                </th>
            <th rowspan="2">DESEMPEÑO</th>
          </tr>
          <tr>
            <th>J</th><th>SJ</th><th>TOTAL</th>
            ${periodSubHeaders}
          </tr>
        </thead>
        <tbody>
          ${rowsHTML}
        </tbody>
      </table>

      <div class="b-footgrid">
        <div class="b-bigline">
          <div class="b-row"><b>TIPO DE FALTAS</b><span class="b-muted">I &nbsp;&nbsp; II &nbsp;&nbsp; III </span></div>
          <div class="b-row"><b>TOTAL IHS</b><span class="b-muted">${totalIHS}</span></div>
          <div class="b-row"><b>INASISTENCIA (J / SJ / TOTAL)</b><span class="b-muted">${sumJ} / ${sumSJ} / ${sumTotal}</span></div>

          <div class="b-row">
            <b>PROMEDIOS</b>
            <span class="b-muted">
              ${promPeriodos.map(p=> isNum(p)?fmt(p,2):"0,00").join("  ")}
              &nbsp;&nbsp; | &nbsp;&nbsp;
              <b>ANUAL:</b> ${isNum(promAnual)?fmt(promAnual,2):"0,00"} &nbsp;&nbsp; ${escapeHTML(desGlobal||"BAJO")}
            </span>
          </div>
        </div>

        <div class="b-legend">
          <h3>Descripciones generales de desempeño (Decreto 1290 de 2009)</h3>
          <div class="lrow"><b>1,00 - 2,99</b><span>DESEMPEÑO BAJO: No supera los desempeños previstos en el área.</span></div>
          <div class="lrow"><b>3,00 - 3,99</b><span>DESEMPEÑO BÁSICO: Logra lo mínimo; debe mejorar para alcanzar niveles esperados.</span></div>
          <div class="lrow"><b>4,00 - 4,59</b><span>DESEMPEÑO ALTO: Alcanza la mayoría de desempeños; buen nivel de desarrollo.</span></div>
          <div class="lrow"><b>4,60 - 5,00</b><span>DESEMPEÑO SUPERIOR: Alcanza la totalidad; cumple de manera cabal.</span></div>
          <div class="b-note b-muted">X: presentó mejoramiento · X: debió presentarlo y no lo hizo</div>
        </div>
      </div>

      <div class="b-footgrid">
        <div class="b-bigline">
          <div class="b-row"><b>PROMOVIDO:</b><span>${escapeHTML(promBo)} &nbsp;&nbsp; <b>  </b> &nbsp;&nbsp; <b>AL GRADO:</b> ${escapeHTML(alGrBo)}</span></div>
          <div class="b-note"><b>ACTIVIDADES DE MEJORAMIENTO:</b><div class="b-muted" style="margin-top:6px;white-space:pre-wrap">${escapeHTML(db.settings.actMej||"")}</div></div>
        </div>
        <div class="b-bigline">
          <div class="b-note"><b>OBSERVACIONES:</b><div class="b-muted" style="margin-top:6px;white-space:pre-wrap">${escapeHTML(db.settings.obsBoletin||"")}</div></div>
        </div>
      </div>

      ${firmasHTML()}
    </div>
  `;
  wrap.innerHTML = html;
}

/* =========================
   CONFIG
   ========================= */
function setPeriods(){
  const val = Number(document.getElementById("setPeriods").value);
  if(val!==2 && val!==3 && val!==4) return;

  db.settings.periods = val;
  db.settings.periodCfg = normalizePeriodCfg(db.settings.periodCfg, val);

  
  db.settings.periodDates = normalizePeriodDates(db.settings.periodDates, val);
saveDB();
  fillPeriodSelects(); fillPeriodSelectsGenerator();
  renderGradebook(); renderFinalArea(); renderGenerator();
  renderPeriodCfgTable();
}
function saveSettings(){
  db.settings.gradeGlobal = (document.getElementById("setGradeGlobal").value||"").trim();
  db.settings.institution = (document.getElementById("setInstitution").value||"").trim();
  const munEl = document.getElementById("setMunicipio"); if(munEl) db.settings.municipio = (munEl.value||"").trim();
  const depEl = document.getElementById("setDepto"); if(depEl) db.settings.depto = (depEl.value||"").trim();
  db.settings.place = (document.getElementById("setPlace").value||"").trim();
  // Si no se diligencia la línea inferior, usar "MUNICIPIO DEPTO" (opcional)
  if(!db.settings.place && (db.settings.municipio || db.settings.depto)){
    db.settings.place = `${db.settings.municipio||""} ${db.settings.depto||""}`.trim();
  }

  const ien = document.getElementById("setIEName"); if(ien) db.settings.ieName = (ien.value||"").trim();
  const nit = document.getElementById("setNIT"); if(nit) db.settings.nit = (nit.value||"").trim();
  const dane = document.getElementById("setDANE"); if(dane) db.settings.dane = (dane.value||"").trim();
  const rn = document.getElementById("setRectorName"); if(rn) db.settings.rectorName = (rn.value||"").trim();
  const dn = document.getElementById("setDirName"); if(dn) db.settings.dirName = (dn.value||"").trim();

  saveDB();
  renderGenerator();
}

/* =========================
   INTEGRIDAD
   ========================= */
function runIntegrity(){
  const out = [];
  const areaIds = new Set(db.areas.map(a=>a.id));
  const studentIds = new Set(db.students.map(s=>s.id));

  const orphanAreas = Object.keys(db.grades).filter(aId=>!areaIds.has(aId));
  out.push(orphanAreas.length ? `• Áreas inexistentes en grades: ${orphanAreas.length}` : "• OK: no hay áreas inexistentes en grades.");

  let ghostStudents=0;
  for(const aId in db.grades){
    const byP = db.grades[aId] || {};
    for(const p in byP){
      const byS = byP[p] || {};
      for(const stId in byS){
        if(!studentIds.has(stId)) ghostStudents++;
      }
    }
  }
  out.push(`• Estudiantes inexistentes referenciados en grades: ${ghostStudents}`);
  out.push(`• Estudiantes: ${db.students.length}`);
  out.push(`• Áreas: ${db.areas.length}`);

  document.getElementById("integrityOut").textContent = out.join("\n");
}

/* =========================
   BUSCADOR RÁPIDO
   ========================= */
function quickSearchStudent(){
  const q = (document.getElementById("quickStudent").value||"").toLowerCase().trim();
  const el = document.getElementById("quickStudentResult");
  if(!q){ el.textContent="—"; return; }
  const s = db.students.find(x => (x.name||"").toLowerCase().includes(q) || (x.ident||"").toLowerCase().includes(q));
  el.textContent = s ? `${s.name}${s.ident?(" · "+s.ident):""}${s.grade?(" · "+s.grade):""}` : "No encontrado";
}
function quickSearchArea(){
  const q = (document.getElementById("quickArea").value||"").toLowerCase().trim();
  const el = document.getElementById("quickAreaResult");
  if(!q){ el.textContent="—"; return; }
  const a = db.areas.find(x => (x.name||"").toLowerCase().includes(q));
  el.textContent = a ? `${a.name}${a.teacher?(" · "+a.teacher):""}${a.ihs?(" · IHS "+a.ihs):""}` : "No encontrado";
}

/* =========================
   XLSX EXPORT (SheetJS)
   ========================= */
function ensureXLSX(){
  if(typeof XLSX === "undefined"){
    alert("No se cargó la librería XLSX. Revisa conexión (CDN) o usa impresión.");
    return false;
  }
  return true;
}
function downloadXLSX(wb, filename){ XLSX.writeFile(wb, filename); }

function exportListadoEstudiantesXLSX(){
  if(!ensureXLSX()) return;
  const aoa = [];
  aoa.push(["Nº","APELLIDOS Y NOMBRES","IDENTIFICACION","GRADO","CÓDIGO"]);
  db.students.forEach((s,i)=> aoa.push([i+1, s.name||"", s.ident||"", s.grade||"", s.code||""]));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
  downloadXLSX(wb, "LISTADO ESTUDIANTES.xlsx");
}

function exportPlanillaActualXLSX(){
  const areaId = document.getElementById("gbArea").value;
  const period = Number(document.getElementById("gbPeriod").value);
  if(!areaId || !period) return alert("Selecciona área y periodo.");
  exportAreaPlanillaXLSX(areaId, period);
}
function exportAreaPlanillaXLSX(areaId, period){
  if(!ensureXLSX()) return;

  const a = db.areas.find(x=>x.id===areaId);
  if(!a) return alert("Área inválida.");

  ensureGrades(areaId, period);
  const cfg = getPeriodCfg(period);

  const headers = [
    "Nº","APELLIDOS Y NOMBRES",
    "J","SJ","TOTAL",
    ...Array.from({length:cfg.prNotes}, (_,i)=>`N${i+1}`),
    "AE","ST","REC.",
    ...Array.from({length:cfg.coNotes}, (_,i)=>`N${i+1}`),
    "AE","ST","REC.",
    "DOC.","EST.","ST","0.1","REC.",
    "NOTA \nDEFINITIVA","ESCALA \nNACIONAL"
  ];
  const aoa = [headers];

  db.students.forEach((s,i)=>{
    const row = (db.grades?.[areaId]?.[period]?.[s.id]) || emptyRow();
    calcRowExcelLike(row, cfg);
    const total = (toNum(row.j)||0) + (toNum(row.sj)||0);
    const nota = row._computed?.nota ?? null;

    const prVals = Array.from({length:cfg.prNotes}, (_,k)=> row[noteField("pr",k+1)] ?? "");
    const coVals = Array.from({length:cfg.coNotes}, (_,k)=> row[noteField("co",k+1)] ?? "");

    aoa.push([
      i+1, s.name||"",
      row.j||0, row.sj||0, total,
      ...prVals,
      row.pr_ae??"", isNum(row._computed?.pr_st)? Number(row._computed.pr_st.toFixed(2)):"", row.pr_rec??"",
      ...coVals,
      row.co_ae??"", isNum(row._computed?.co_st)? Number(row._computed.co_st.toFixed(2)):"", row.co_rec??"",
      row.ac_doc??"", row.ac_est??"", isNum(row._computed?.ac_st)? Number(row._computed.ac_st.toFixed(2)):"", isNum(row._computed?.ac_w)? Number(row._computed.ac_w.toFixed(2)):"", row.ac_rec??"",
      isNum(nota)? Number(nota.toFixed(2)) : "",
      isNum(nota)? escalaNacional_D(nota) : " "
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `P${period}`);
  downloadXLSX(wb, `AREAS_${(a.name||"AREA").replace(/[^\w]+/g,"_")}_P${period}.xlsx`);
}

function exportFinalAreaActualXLSX(){
  const areaId = document.getElementById("faArea").value;
  if(!areaId) return alert("Selecciona un área.");
  exportFinalAreaXLSX(areaId);
}
function exportFinalAreaXLSX(areaId){
  if(!ensureXLSX()) return;
  const a = db.areas.find(x=>x.id===areaId);
  if(!a) return alert("Área inválida.");

  const periods = db.settings.periods;
  const headers = ["Nº","APELLIDOS Y NOMBRES","J","SJ","TOTAL"];
  for(let p=1;p<=periods;p++) headers.push("NOTA","DESEMPEÑO");
  headers.push("NOTA \nDEFINITIVA","ESCALA\nNACIONAL");

  const aoa = [headers];
  db.students.forEach((s,i)=>{
    const fin = computeAreaFinalForStudent(areaId, s.id);
    const row = [i+1, s.name||"", fin.inaJ, fin.inaSJ, fin.inaTotal];
    fin.per.forEach(x=>{
      row.push(isNum(x.nota)? Number(x.nota.toFixed(2)) : "");
      row.push(x.des || " ");
    });
    row.push(isNum(fin.notaDef)? Number(fin.notaDef.toFixed(2)) : "");
    row.push(fin.desDef || " ");
    aoa.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "FINAL");
  downloadXLSX(wb, `FINAL_AREA_${(a.name||"AREA").replace(/[^\w]+/g,"_")}.xlsx`);
}

function exportGeneradorActualXLSX(){
  const studentId = document.getElementById("genStudent").value;
  if(!studentId) return alert("Selecciona un estudiante.");
  exportBoletinPDF(studentId);
}
function exportGeneradorFinalXLSX(studentId){
  if(!ensureXLSX()) return;
  const st = db.students.find(s=>s.id===studentId);
  if(!st) return alert("Estudiante inválido.");

  const periods = db.settings.periods;
  const headers = ["AREA Y/O \nASIGNATURA","IHS","DOCENTE","J","SJ","TOTAL"];
  for(let p=1;p<=periods;p++) headers.push("NOTA","DESEMPEÑO","MEJ.");
  headers.push("DEFINITIVA ANUAL","DESEMPEÑO");

  const aoa = [headers];
  db.areas.forEach(a=>{
    const fin = computeAreaFinalForStudent(a.id, studentId);
    const anual = avgSimple(fin.per.map(x=>x.nota));
    const desAnual = isNum(anual) ? escalaNacional_SIN_D(anual) : " ";

    const row = [a.name||"", a.ihs||"", a.teacher||"", fin.inaJ, fin.inaSJ, fin.inaTotal];
    fin.per.forEach(x=>{
      row.push(isNum(x.nota)? Number(x.nota.toFixed(2)) : "");
      row.push(isNum(x.nota)? escalaNacional_SIN_D(x.nota) : " ");
      row.push("");
    });
    row.push(isNum(anual)? Number(anual.toFixed(2)) : "");
    row.push(desAnual);
    aoa.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "FINAL");
  const name = (st.name||"ESTUDIANTE").replace(/[^\w]+/g,"_");
  downloadXLSX(wb, `GENERADOR_FINAL_${name}.xlsx`);
}

/* =========================
   BOLETÍN (EXPORT PDF + XLSX) — Formato tipo "Boletin.pdf" (carta)
   - Se ejecuta desde el botón: "Exportar GENERADOR_FINAL_...xlsx"
   ========================= */
function ensurePDFLibs(){
  // jsPDF UMD expone window.jspdf.jsPDF (CDN) y algunas variantes exponen window.jsPDF
  const hasCanvas = (typeof html2canvas !== "undefined");
  const hasJsPDF = !!((window.jspdf && window.jspdf.jsPDF) || window.jsPDF || (typeof jspdf !== "undefined" && jspdf.jsPDF));
  return hasCanvas && hasJsPDF;
}
function ensureExcelJS(){
  // ExcelJS debe exponer constructor Workbook en el global
  return (typeof ExcelJS !== "undefined") && (typeof ExcelJS.Workbook === "function");
}
function downloadBlob(blob, filename){
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 2000);
}
function safeFileName(name){
  return String(name||"ESTUDIANTE").trim().replace(/[^\w]+/g,"_").replace(/^_+|_+$/g,"") || "ESTUDIANTE";
}
function toast(msg, ms=2200){
  try{
    msg = String(msg ?? "");
    ms = Number(ms) || 2200;

    // Reusar (o crear) un contenedor mínimo, sin tocar estilos globales.
    let el = document.getElementById("toast");
    if(!el){
      el = document.createElement("div");
      el.id = "toast";
      el.style.position = "fixed";
      el.style.right = "14px";
      el.style.bottom = "14px";
      el.style.zIndex = "99999";
      el.style.background = "rgba(0,0,0,0.82)";
      el.style.color = "#fff";
      el.style.padding = "10px 12px";
      el.style.borderRadius = "10px";
      el.style.fontSize = "14px";
      el.style.maxWidth = "70vw";
      el.style.boxShadow = "0 8px 24px rgba(0,0,0,0.25)";
      el.style.opacity = "0";
      el.style.transition = "opacity .15s ease";
      document.body.appendChild(el);
    }

    el.textContent = msg;
    el.style.opacity = "1";

    // timeout por elemento (no global)
    if(el.__toastTimer) clearTimeout(el.__toastTimer);
    el.__toastTimer = setTimeout(()=>{ el.style.opacity = "0"; }, ms);
  }catch(e){
    // fallback silencioso
    try{ console.log(msg); }catch(_){}
  }
}

function normalizeGradeToken(grade){
  let g = String(grade||"").trim();
  if(!g) return "NA";
  g = g.replace(/grado/ig, "").trim();
  g = safeFileName(g);
  return (g || "NA").toUpperCase();
}
function fmtBoletin(n){
  return isNum(n) ? n.toFixed(2).replace(".", ",") : "";
}
function ordinalPeriodo(p){
  if(p===1) return "PRIMER PERIODO";
  if(p===2) return "SEGUNDO PERIODO";
  if(p===3) return "TERCER PERIODO";
  return `PERIODO ${p}`;
}
function computeBoletinData(studentId){
  const st = db.students.find(s=>s.id===studentId);
  if(!st) return null;

  const periods = db.settings.periods || 3;

  const ieName = (db.settings.ieName || db.settings.institution || "").trim();
  const resol = (db.settings.institution || "").trim(); // texto libre (ej: Resolución...)
  const nit = (db.settings.nit || "").trim();
  const dane = (db.settings.dane || "").trim();
  const place = (db.settings.place || "").trim();
  const shield = db.settings.images?.shield || "";

  const year = (db.settings.year || new Date().getFullYear()).toString();
  const jornada = (db.settings.jornada || "MAÑANA").trim();
  const grupo = (db.settings.grupo || "").trim();
  const fecha = (db.settings.fechaBoletin || "").trim();

  const grade = (st.grade || db.settings.gradeGlobal || "").trim();

  const items = db.areas.map(a=>{
    const fin = computeAreaFinalForStudent(a.id, studentId);
    const per = fin.per.map(x=>{
      const nota = x.nota;
      const des = isNum(nota) ? escalaNacional_SIN_D(nota) : "";
      return { nota, des };
    });
    const anualNota = fin.notaDef;
    const anualDes = isNum(anualNota) ? escalaNacional_SIN_D(anualNota) : "";
    const ihsNum = toNum(a.ihs);
    return {
      area: (a.name||"").trim(),
      ihs: isNum(ihsNum)? ihsNum : (a.ihs||""),
      docente: (a.teacher||"").trim(),
      inaJ: fin.inaJ||0,
      inaSJ: fin.inaSJ||0,
      inaTotal: fin.inaTotal||0,
      per,
      anualNota,
      anualDes
    };
  });

  const totalIHS = items.reduce((acc,it)=> acc + (Number(it.ihs)||0), 0);

  // Promedios por periodo (promedio simple de notas válidas > 0,9, como en el resto del sistema)
  const promPeriodos = [];
  for(let p=0;p<periods;p++){
    const notas = items.map(it=> it.per[p]?.nota ?? null);
    promPeriodos.push(avgIfGT09_periodNotes(notas));
  }
  const promAnual = avgIfGT09_periodNotes(promPeriodos);
  const promAnualDes = isNum(promAnual) ? escalaNacional_SIN_D(promAnual) : "BAJO";

  const actMej = (db.settings.actMej || "").trim();
  const obs = (db.settings.obsBoletin || "").trim();
  const promovido = (db.settings.promovido || "SI").trim();
  const alGrado = (db.settings.alGrado || "").trim();
  const posGrupo = ""; // (no existe campo en el sistema; se deja en blanco como en el PDF)

  const dirSign = db.settings.images?.dirSign || "";
  const dirSeal = db.settings.images?.dirSeal || "";
  const dirName = (db.settings.dirName || "").trim();

  return { st, periods, ieName, resol, nit, dane, place, shield, year, jornada, grupo, fecha, grade,
           items, totalIHS, promPeriodos, promAnual, promAnualDes,
           actMej, obs, promovido, alGrado, posGrupo,
           dirSign, dirSeal, dirName };
}
function buildBoletinExportNode(data){
  const root = document.createElement("div");
  root.id = "boletinExportRoot";
  root.style.background = "#fff";
  root.style.color = "#000";

  const logoHTML = data.shield ? `<img class="logo" src="${data.shield}" alt="Escudo" />` : `<div class="logo ph"></div>`;

  const headerLeft = `
    <div class="inst">
      <div class="ie">${escapeHTML(data.ieName || "INSTITUCIÓN EDUCATIVA")}</div>
      <div class="small">${escapeHTML(data.resol || "")}</div>
      <div class="small">${data.nit ? `NIT ${escapeHTML(data.nit)}` : ""}</div>
      <div class="small">${data.dane ? `CODIGO DANE ${escapeHTML(data.dane)}` : ""}</div>
      <div class="small">${escapeHTML(data.place || "")}</div>
    </div>
  `;

  const headerRight = `
    <div class="title">INFORME DE VALORACION DEL DESEMPEÑO</div>
    <table class="mini">
      <tr>
        <td class="k">GRADO:</td><td>${escapeHTML(data.grade||"")}</td>
        <td class="k">GRUPO:</td><td>${escapeHTML(data.grupo||"")}</td>
        <td class="k">FECHA</td><td>${escapeHTML(data.fecha||"")}</td>
      </tr>
      <tr>
        <td class="k">ESTUDIANTE:</td><td colspan="5">${escapeHTML(data.st.name||"")}</td>
      </tr>
      <tr>
        <td class="k">IDENTIFICACION:</td><td colspan="5"><b>${escapeHTML(data.st.ident||"")}</b></td>
      </tr>
      <tr>
        <td class="k">AÑO LECTIVO:</td><td>${escapeHTML(data.year||"")}</td>
        <td class="k">JORNADA:</td><td colspan="3">${escapeHTML(data.jornada||"")}</td>
      </tr>
    </table>
  `;

  const headPeriods = Array.from({length:data.periods}, (_,i)=>{
    const p = i+1;
    return `<th colspan="2">${ordinalPeriodo(p)}</th><th rowspan="2" class="mej">MEJ.</th>`;
  }).join("");

  const headPeriodCols = Array.from({length:data.periods}, ()=> `<th>NOTA</th><th>DESEMPEÑO</th>`).join("");

  const bodyRows = data.items.map(it=>{
    const perCells = it.per.map((x)=>`
      <td class="c">${fmtBoletin(x.nota)}</td>
      <td class="c">${escapeHTML(x.des||"")}</td>
      <td class="c"></td>
    `).join("");

    return `
      <tr>
        <td class="area">${escapeHTML(it.area)}</td>
        <td class="c">${escapeHTML(String(it.ihs??""))}</td>
        <td class="doc">${escapeHTML(it.docente||"")}</td>
        <td class="c">${it.inaJ||0}</td>
        <td class="c">${it.inaSJ||0}</td>
        <td class="c">${it.inaTotal||0}</td>
        ${perCells}
        <td class="c">${fmtBoletin(it.anualNota)}</td>
        <td class="c"><b>${escapeHTML(it.anualDes||"")}</b></td>
      </tr>
    `;
  }).join("");

  const conductaRow = `
    <tr class="thick-top">
      <td colspan="3" class="area"><b>CONDUCTA Y DISCIPLINA</b></td>
      <td colspan="3" class="c"><b>TIPO DE FALTAS</b> &nbsp; I0 &nbsp;&nbsp; II0 &nbsp;&nbsp; III0</td>
      ${Array.from({length:data.periods}, ()=>`
        <td class="c">${fmtBoletin(0)}</td>
        <td class="c"></td>
        <td class="c"></td>
      `).join("")}
      <td class="c"><b>${fmtBoletin(0)}</b></td>
      <td class="c"><b>BAJO</b></td>
    </tr>
  `;

  const promsCells = data.promPeriodos.map(n=>`
    <td class="c">${fmtBoletin(n)}</td>
    <td class="c"></td>
    <td class="c"></td>
  `).join("");

  const promsRow = `
    <tr>
      <td class="c"><b>TOTAL IHS</b></td>
      <td class="c"><b>${data.totalIHS||0}</b></td>
      <td colspan="4" class="c"><b>PROMEDIOS</b></td>
      ${promsCells}
      <td class="c"><b>${fmtBoletin(data.promAnual)}</b></td>
      <td class="c"><b>${escapeHTML(data.promAnualDes||"BAJO")}</b></td>
    </tr>
  `;

  const promoTable = `
    <table class="promo">
      <tr>
        <td class="k">PROMOVIDO:</td>
        <td class="c ${data.promovido==="SI"?"sel":""}">SI</td>
        <td class="c ${data.promovido==="NO"?"sel":""}">NO</td>
        <td class="k">AL GRADO:</td>
        <td class="fill">${escapeHTML(data.alGrado||"")}</td>
        <td class="k">POSICION GRUPO:</td>
        <td class="fill">${escapeHTML(data.posGrupo||"")}</td>
      </tr>
    </table>
  `;

  const obsLines = Array.from({length:5}, ()=>`<div class="line"></div>`).join("");

  const signHTML = (data.dirSign || data.dirSeal) ? `
    <div class="sign">
      ${data.dirSign ? `<img src="${data.dirSign}" alt="Firma docente" />` : ""}
      ${data.dirSeal ? `<img src="${data.dirSeal}" alt="Sello docente" />` : ""}
      <div class="sname">${escapeHTML(data.dirName||"DOCENTE")}</div>
    </div>
  ` : "";

  root.innerHTML = `
    <style>
      /* Estilos locales para exportación (no afectan la UI) */
      #boletinExportRoot{ width: 190mm; padding: 0; margin: 0 auto; font-family: Arial, Helvetica, sans-serif; }
      #boletinExportRoot table{ border-collapse: collapse; width:100%; }
      #boletinExportRoot .box{ border:2px solid #000; }
      #boletinExportRoot .hdr{ table-layout: fixed; }
      #boletinExportRoot .hdr td{ border:2px solid #000; vertical-align: middle; }
      #boletinExportRoot .logoCell{ width: 30mm; text-align:center; padding:4mm; }
      #boletinExportRoot .logo{ width: 22mm; height: 22mm; object-fit: contain; display:block; margin:0 auto; }
      #boletinExportRoot .logo.ph{ border:1px solid #000; width:22mm; height:22mm; }
      #boletinExportRoot .instCell{ width: 80mm; padding:3mm 4mm; }
      #boletinExportRoot .inst .ie{ font-weight:700; font-size: 11pt; text-align:left; }
      #boletinExportRoot .inst .small{ font-size: 8.5pt; line-height: 1.2; }
      #boletinExportRoot .infoCell{ padding:0; }
      #boletinExportRoot .title{ font-weight:700; text-align:center; font-size: 10pt; padding:2mm 0; border-bottom:2px solid #000; }
      #boletinExportRoot .mini{ width:100%; font-size: 8.5pt; }
      #boletinExportRoot .mini td{ border:1px solid #000; padding:1.2mm 1.5mm; }
      #boletinExportRoot .mini td.k{ font-weight:700; width: 18mm; white-space:nowrap; }
      #boletinExportRoot .tbl{ margin-top: 6mm; font-size: 8.5pt; table-layout: fixed; }
      #boletinExportRoot .tbl th, #boletinExportRoot .tbl td{ border:1px solid #000; padding:1.2mm 1.2mm; }
      #boletinExportRoot .tbl thead th{ font-weight:700; text-align:center; }
      #boletinExportRoot .tbl .area{ text-align:left; width: 36mm; }
      #boletinExportRoot .tbl .doc{ text-align:left; width: 26mm; }
      #boletinExportRoot .tbl .c{ text-align:center; }
      #boletinExportRoot .tbl .mej{ width: 5mm; writing-mode: vertical-rl; transform: rotate(180deg); font-weight:700; letter-spacing: 0.5px; }
      #boletinExportRoot .tbl .thick-top td{ border-top:2px solid #000; }
      #boletinExportRoot .promo{ margin-top: 2mm; font-size: 8.5pt; }
      #boletinExportRoot .promo td{ border:2px solid #000; padding:1.2mm 1.5mm; }
      #boletinExportRoot .promo td.k{ font-weight:700; white-space:nowrap; width: 24mm; }
      #boletinExportRoot .promo td.c{ text-align:center; width: 10mm; }
      #boletinExportRoot .promo td.fill{ width:auto; }
      #boletinExportRoot .promo td.sel{ font-weight:700; }
      #boletinExportRoot .notes{ margin-top: 8mm; display:grid; grid-template-columns: 1fr 1.7fr; gap: 8mm; }
      #boletinExportRoot .notes .nbox{ border:2px solid #000; padding:3mm; min-height: 24mm; font-size: 8.5pt; }
      #boletinExportRoot .notes .nbox h4{ margin:0 0 2mm 0; font-size: 9pt; }
      #boletinExportRoot .notes .line{ border-bottom:1px solid #000; height: 6mm; }
      #boletinExportRoot .legend{ margin-top: 8mm; border:2px solid #000; }
      #boletinExportRoot .legend th, #boletinExportRoot .legend td{ border:1px solid #000; padding:1.6mm; font-size: 8.2pt; }
      #boletinExportRoot .legend th{ text-align:center; font-weight:700; }
      #boletinExportRoot .sign{ position: relative; margin-top: 6mm; width: 65mm; margin-left:auto; text-align:center; }
      #boletinExportRoot .sign img{ max-width: 65mm; max-height: 22mm; object-fit: contain; display:block; margin: 0 auto; }
      #boletinExportRoot .sname{ font-weight:700; font-size: 9pt; margin-top: 1mm; }
    </style>

    <table class="hdr">
      <tr>
        <td class="logoCell">${logoHTML}</td>
        <td class="instCell">${headerLeft}</td>
        <td class="infoCell">${headerRight}</td>
      </tr>
    </table>

    <table class="tbl">
      <thead>
        <tr>
          <th rowspan="3">AREA Y/O<br>ASIGNATURA</th>
          <th rowspan="3">IHS</th>
          <th rowspan="3">DOCENTE</th>
          <th colspan="3" rowspan="2">INASISTENCIA</th>
          <th colspan="${data.periods*3}">PERIODOS</th>
          <th colspan="2" rowspan="2">DEFINITIVA ANUAL</th>
        </tr>
        <tr>
          ${headPeriods}
        </tr>
        <tr>
          <th>J</th><th>SJ</th><th>TOTAL</th>
          ${headPeriodCols}
        </tr>
      </thead>
      <tbody>
        ${bodyRows}
        ${conductaRow}
        ${promsRow}
      </tbody>
    </table>

    ${promoTable}

    <div class="notes">
      <div class="nbox">
        <h4>ACTIVIDADES DE MEJORAMIENTO:</h4>
        <div><b>M:</b> El(a) estudiante presentó actividades de mejoramiento en el área / asignatura</div>
        <div style="margin-top:2mm"><b>N:</b> El(a) estudiante debió presentar actividades de mejoramiento en el área / asignatura y no lo hizo</div>
      </div>
      <div class="nbox">
        <h4>OBSERVACIONES:</h4>
        <div>${escapeHTML(data.obs||"")}</div>
        ${obsLines}
      </div>
    </div>

    <table class="legend" style="margin-top:8mm">
      <tr><th colspan="4">DESCRIPCIONES GENERALES DE DESEMPEÑO (Decreto 1290 de abril 16 de 2009)</th></tr>
      <tr>
        <td style="width:28mm"><b>DESEMPEÑO BAJO</b></td>
        <td style="width:22mm" class="c"><b>1,00 - 2,99</b></td>
        <td colspan="2">No supera los desempeños previstos en el área, teniendo bajo rendimiento en todos los procesos de aprendizaje por lo que no alcanza objetivos y metas establecidas.</td>
      </tr>
      <tr>
        <td><b>DESEMPEÑO BASICO</b></td>
        <td class="c"><b>3,00 - 3,99</b></td>
        <td colspan="2">Logra lo mínimo en los procesos de formación y aunque puede ser promovido en su proceso académico, debe mejorar su desempeño para alcanzar los niveles de aprendizaje esperados.</td>
      </tr>
      <tr>
        <td><b>DESEMPEÑO ALTO</b></td>
        <td class="c"><b>4,00 - 4,59</b></td>
        <td colspan="2">Alcanza la mayoría de los desempeños previstos en el área de formación, demostrando un buen nivel de desarrollo.</td>
      </tr>
      <tr>
        <td><b>DESEMPEÑO SUPERIOR</b></td>
        <td class="c"><b>4,60 - 5,00</b></td>
        <td colspan="2">Alcanza la totalidad de los desempeños previstos en el área de formación, cumple de manera cabal con todos los procesos de desarrollo integral superando los objetivos y metas.</td>
      </tr>
    </table>

    ${signHTML}
  `;

  return root;
}

async function exportBoletinPDF(studentId){
  const data = computeBoletinData(studentId);
  if(!data) return alert("Estudiante inválido.");

  if(!ensurePDFLibs()){
    alert("No se cargaron librerías de PDF (html2canvas / jsPDF). Verifica conexión a internet o recarga la página.");
    return;
  }

  // Crear nodo aislado para exportar (evita imprimir cabecera oscura del app)
  const host = document.createElement("div");
  host.style.position = "fixed";
  host.style.left = "-99999px";
  host.style.top = "0";
  host.style.width = "1000px";
  host.style.background = "#fff";
  host.style.padding = "0";
  host.style.margin = "0";
  document.body.appendChild(host);

  const node = buildBoletinRoot(data);
  host.appendChild(node);

  try{

  // Render con alta nitidez pero tamaño de archivo controlado
  const dpr = window.devicePixelRatio || 1;
  const scale = getPDFRenderScale(dpr);
  let canvas = await html2canvas(node, { backgroundColor:"#ffffff", scale, useCORS: true, logging:false });
  // Limitar a aprox. 300 DPI en carta (reduce peso sin perder legibilidad)
  canvas = downscaleCanvasIfNeeded(canvas, 2550, 3300);
  const imgData = canvas.toDataURL("image/jpeg", 0.92);

  const jsPDFCtor = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF || (typeof jspdf !== "undefined" && jspdf.jsPDF);
  if(!jsPDFCtor){ alert("No se pudo inicializar jsPDF. Verifica conexión o recarga la página."); host.remove(); return; }

  const pdf = new jsPDFCtor({ orientation:"p", unit:"pt", format:"letter", compress:true });

  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = 18; // ~6mm
  const maxW = pageW - margin*2;
  const maxH = pageH - margin*2;

  const imgW = canvas.width;
  const imgH = canvas.height;

  // Forzar a una sola página tamaño carta (ajusta por ancho y alto)
  const fitScale = Math.min(maxW / imgW, maxH / imgH);
  const renderW = imgW * fitScale;
  const renderH = imgH * fitScale;

  // Centrar dentro de márgenes (sin salirse)
  const x = Math.max(margin, (pageW - renderW) / 2);
  const y = Math.max(margin, (pageH - renderH) / 2);

  pdf.addImage(imgData, "JPEG", x, y, renderW, renderH, undefined, "FAST");

  // Nombre por defecto: NOMBRES_Y_APELLIDOS_GRADO_PERIODO.pdf
  // (incluye el valor del grado y el periodo seleccionado; en modo FINAL usa "FINAL")
  const name = safeFileName(data.st.name).toUpperCase();
  const gradeTok = normalizeGradeToken(data.grade || data.st?.grade || db.settings.gradeGlobal || "");
  const mode = document.getElementById("genMode")?.value || "final";
  const perSel = Number(document.getElementById("genPeriod")?.value || 1) || 1;
  const perTok = (mode === "period") ? String(perSel) : "FINAL";
  pdf.save(`${name}_${gradeTok}_${perTok}.pdf`);

  }catch(err){
    console.error(err);
    alert("No se pudo exportar PDF. Usa \"Imprimir PDF\".");
  }

  host.remove();
}


async function exportBoletinXLSX(studentId){
  const data = computeBoletinData(studentId);
  if(!data) return alert("Estudiante inválido.");

  // Si ExcelJS no carga, usar fallback (SheetJS) para no bloquear al usuario
  if(!ensureExcelJS()){
    exportGeneradorFinalXLSX(studentId);
    return;
  }

  try{
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("BOLETIN", {
    pageSetup: {
      paperSize: 1, // letter
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      margins: { left:0.3, right:0.3, top:0.4, bottom:0.4, header:0, footer:0 }
    }
  });

  const periods = Number(data.periods || 4);
  const totalCols = 6 + periods*3 + 2;
  const periodStartCol = 7;
  const periodEndCol = 6 + periods*3;
  const annualStartCol = periodEndCol + 1;
  const annualEndCol = annualStartCol + 1;

  // Column widths dinámicas (base + 3 por periodo + anual)
  const colW = [22,6,14,4,4,6];
  for(let p=0; p<periods; p++) colW.push(7,12,3);
  colW.push(7,12);
  ws.columns = colW.map(w=>({ width:w }));

  const thin = { style:"thin", color:{ argb:"FF000000" } };
  const borderThin = { top:thin, left:thin, bottom:thin, right:thin };
  const center = { vertical:"middle", horizontal:"center", wrapText:true };
  const left = { vertical:"middle", horizontal:"left", wrapText:true };

  // Helper to set border/alignment
  function box(r,c,align){
    ws.getCell(r,c).border = borderThin;
    ws.getCell(r,c).alignment = align || center;
    ws.getCell(r,c).font = ws.getCell(r,c).font || { size:9 };
  }

  // Título (rows 1-3)
  ws.mergeCells(1,1,1,totalCols);
  ws.getCell(1,1).value = (data.schoolName || "").toUpperCase();
  ws.getCell(1,1).font = { bold:true, size:14 };
  ws.getCell(1,1).alignment = { vertical:"middle", horizontal:"center" };

  ws.mergeCells(2,1,2,totalCols);
  ws.getCell(2,1).value = "BOLETÍN FINAL DE EVALUACIÓN";
  ws.getCell(2,1).font = { bold:true, size:12 };
  ws.getCell(2,1).alignment = { vertical:"middle", horizontal:"center" };

  ws.mergeCells(3,1,3,totalCols);
  ws.getCell(3,1).value = `AÑO LECTIVO: ${data.year || ""}`;
  ws.getCell(3,1).font = { bold:true, size:10 };
  ws.getCell(3,1).alignment = { vertical:"middle", horizontal:"center" };

  // Datos estudiante (rows 4-6)
  const info = [
    ["ESTUDIANTE:", data.st?.name || "", "GRADO:", data.grade || ""],
    ["IDENTIFICACIÓN:", data.st?.ident || "", "JORNADA:", data.jornada || ""],
    ["NIT:", data.nit || "", "DANE:", data.dane || ""],
  ];

  let row = 4;
  info.forEach(line=>{
    // Distribución: (1-2) label/value, (3-4) label/value y resto vacío
    ws.getCell(row,1).value = line[0]; ws.getCell(row,1).font={bold:true, size:9};
    ws.mergeCells(row,2,row,Math.min(6,totalCols)); ws.getCell(row,2).value = line[1];
    ws.getCell(row,Math.min(7,totalCols)).value = line[2]; ws.getCell(row,Math.min(7,totalCols)).font={bold:true, size:9};
    ws.mergeCells(row,Math.min(8,totalCols),row,totalCols); ws.getCell(row,Math.min(8,totalCols)).value = line[3];

    for(let c=1;c<=totalCols;c++){
      const al = (c===2 || c>=Math.min(8,totalCols)) ? left : center;
      box(row,c,al);
    }
    row++;
  });

  // Escudo (si existe)
  if(data.shield && String(data.shield).startsWith("data:image")){
    try{
      const ext = data.shield.includes("image/png") ? "png" : "jpeg";
      const imgId = wb.addImage({ base64: data.shield, extension: ext });
      ws.addImage(imgId, { tl:{ col:0.2, row:0.4 }, ext:{ width:90, height:90 } });
    }catch(e){ console.warn("No pude insertar escudo en Excel", e); }
  }

  // Tabla principal (header 3 filas)
  const r0 = 7;

  // Row 7
  ws.getCell(r0,1).value="AREA Y/O\nASIGNATURA";
  ws.getCell(r0,2).value="IHS";
  ws.getCell(r0,3).value="DOCENTE";
  ws.mergeCells(r0,4,r0+1,6); ws.getCell(r0,4).value="INASISTENCIA";
  ws.mergeCells(r0,periodStartCol,r0,periodEndCol); ws.getCell(r0,periodStartCol).value="PERIODOS";
  ws.mergeCells(r0,annualStartCol,r0+1,annualEndCol); ws.getCell(r0,annualStartCol).value="DEFINITIVA ANUAL";

  // Row 8: period names + annual subheaders
  let col = periodStartCol;
  for(let p=1;p<=periods;p++){
    ws.mergeCells(r0+1,col,r0+1,col+1);
    ws.getCell(r0+1,col).value = ordinalPeriodo(p);
    ws.getCell(r0+1,col+2).value = "MEJ.";
    col += 3;
  }
  ws.getCell(r0+1,annualStartCol).value="NOTA";
  ws.getCell(r0+1,annualEndCol).value="DESEMPEÑO";

  // Row 9: inasistencia + subheaders por periodo
  ws.getCell(r0+2,4).value="J";
  ws.getCell(r0+2,5).value="SJ";
  ws.getCell(r0+2,6).value="TOTAL";
  col = periodStartCol;
  for(let p=1;p<=periods;p++){
    ws.getCell(r0+2,col).value="NOTA";
    ws.getCell(r0+2,col+1).value="DESEMPEÑO";
    col += 3;
  }

  // Merge vertical headers A,B,C
  ws.mergeCells(r0,1,r0+2,1);
  ws.mergeCells(r0,2,r0+2,2);
  ws.mergeCells(r0,3,r0+2,3);

  // Styling header area
  for(let rr=r0; rr<=r0+2; rr++){
    for(let cc=1; cc<=totalCols; cc++){
      const cell = ws.getCell(rr,cc);
      cell.font = { bold:true, size:9 };
      cell.alignment = center;
      cell.border = borderThin;
      // Header fill suave
      cell.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF2F2F2" } };
    }
  }

  // Body rows
  let r = r0+3;
  data.items.forEach(it=>{
    ws.getCell(r,1).value = it.area;
    ws.getCell(r,2).value = it.ihs || "";
    ws.getCell(r,3).value = it.teacher || "";
    ws.getCell(r,4).value = it.j || "";
    ws.getCell(r,5).value = it.sj || "";
    ws.getCell(r,6).value = it.total || "";

    let c = periodStartCol;
    for(let p=0;p<periods;p++){
      const rec = it.periodos[p] || {};
      const nota = rec.nota;
      ws.getCell(r,c).value = isNum(nota) ? Number(nota.toFixed(2)) : "";
      ws.getCell(r,c+1).value = rec.des || "";
      ws.getCell(r,c+2).value = rec.mej || "";
      c += 3;
    }
    ws.getCell(r,annualStartCol).value = isNum(it.anual) ? Number(it.anual.toFixed(2)) : "";
    ws.getCell(r,annualEndCol).value = it.anualDes || "";

    for(let cc=1; cc<=totalCols; cc++){
      const cell = ws.getCell(r,cc);
      cell.border = borderThin;
      cell.font = { size:9 };
      cell.alignment = (cc===1 || cc===3) ? left : center;
    }
    r++;
  });

  // Promedios
  ws.getCell(r,1).value = "PROMEDIO";
  ws.mergeCells(r,1,r,6);
  ws.getCell(r,1).font={bold:true, size:9};
  ws.getCell(r,1).alignment=center;

  let c = periodStartCol;
  for(let p=0;p<periods;p++){
    const v = data.promPeriodos[p];
    ws.getCell(r,c).value = isNum(v) ? Number(v.toFixed(2)) : "";
    ws.getCell(r,c+1).value = "";
    ws.getCell(r,c+2).value = "";
    c += 3;
  }
  ws.getCell(r,annualStartCol).value = isNum(data.promAnual) ? Number(data.promAnual.toFixed(2)) : "";
  ws.getCell(r,annualEndCol).value = data.promAnualDes || "BAJO";
  ws.getCell(r,annualStartCol).font={bold:true, size:9};
  ws.getCell(r,annualEndCol).font={bold:true, size:9};

  for(let cc=1; cc<=totalCols; cc++){
    ws.getCell(r,cc).border = borderThin;
    ws.getCell(r,cc).alignment = (cc<=6) ? center : center;
    if(cc<=6) ws.getCell(r,cc).fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF2F2F2" } };
  }
  r += 2;

  // Observaciones
  ws.mergeCells(r,1,r,totalCols);
  ws.getCell(r,1).value = `OBSERVACIONES: ${data.obs || ""}`;
  ws.getCell(r,1).alignment = left;
  ws.getCell(r,1).font = { size:9 };
  ws.getCell(r,1).border = borderThin;
  r += 2;

  // Leyenda de desempeño (tabla 4 filas)
  const legendRows = [
    ["DESEMPEÑO BAJO","1,00 - 2,99","Alcanza de forma mínima los desempeños propuestos, requiriendo acompañamiento continuo."],
    ["DESEMPEÑO BÁSICO","3,00 - 3,99","Alcanza los desempeños esenciales previstos para el área de formación."],
    ["DESEMPEÑO ALTO","4,00 - 4,59","Alcanza la mayoría de los desempeños propuestos, demostrando buen nivel de desarrollo."],
    ["DESEMPEÑO SUPERIOR","4,60 - 5,00","Alcanza la totalidad de los desempeños propuestos superando los objetivos y metas."]
  ];

  legendRows.forEach(rowArr=>{
    ws.getCell(r,1).value = rowArr[0]; ws.mergeCells(r,1,r,3);
    ws.getCell(r,4).value = rowArr[1]; ws.mergeCells(r,4,r,5);
    ws.getCell(r,6).value = rowArr[2]; ws.mergeCells(r,6,r,totalCols);

    for(let cc=1; cc<=totalCols; cc++){
      const cell=ws.getCell(r,cc);
      cell.border = borderThin;
      cell.font = { size:9, bold: (cc===1||cc===4) };
      cell.alignment = (cc<=5)? center : left;
      if(cc<=5) cell.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF8F8F8" } };
    }
    r++;
  });

  // Firmas (si existen)
  const rr = r + 1;
  ws.mergeCells(rr,1,rr,totalCols);
  ws.getCell(rr,1).value = `Lugar: ${data.place || ""}    Fecha: ${data.fecha || ""}`;
  ws.getCell(rr,1).font = { size:9 };
  ws.getCell(rr,1).alignment = left;

  // Insertar firma/sello como imagen (si hay)
  if((data.dirSign && String(data.dirSign).startsWith("data:image")) || (data.dirSeal && String(data.dirSeal).startsWith("data:image"))){
    try{
      const imgSrc = data.dirSeal || data.dirSign;
      const ext = imgSrc.includes("image/png") ? "png" : "jpeg";
      const imgId = wb.addImage({ base64: imgSrc, extension: ext });
      // Ubicación hacia el final, sin depender del número exacto de columnas
      ws.addImage(imgId, { tl:{ col: Math.max(0, totalCols-5) + 0.2, row: rr + 0.2 }, ext:{ width: 220, height: 70 } });
    }catch(e){ console.warn("No pude insertar firma/sello en Excel", e); }
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const name = safeFileName(data.st.name);
  downloadBlob(blob, `GENERADOR_FINAL_${name}.xlsx`);
  }catch(err){
    console.error(err);
    console.warn("Fallo ExcelJS, usando exportación XLSX básica (SheetJS).");
    try{ exportGeneradorFinalXLSX(studentId); }catch(e2){ console.error(e2); alert("No se pudo exportar XLSX."); }
  }

}


async function exportBoletinPDFyXLSX(studentId){
  // XLSX primero (rápido), PDF después
  try{ await exportBoletinXLSX(studentId); }catch(e){ console.error(e); alert("No se pudo exportar XLSX."); }
  try{ await exportBoletinPDF(studentId); }catch(e){ console.error(e); alert("No se pudo exportar PDF."); }
}


/* =========================
   HELPERS
   ========================= */
function escapeHTML(str){
  return String(str ?? "")
    .replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
    .replaceAll('"',"&quot;").replaceAll("'","&#039;");
}


// =========================
// PDF RENDER HELPERS (nitidez + tamaño controlado)
// =========================
function getPDFRenderScale(dpr){
  // Suficiente para texto nítido sin disparar el peso del PDF.
  // En pantallas HiDPI (dpr>1) se aprovecha el dpr, pero se limita el máximo.
  return Math.min(3, Math.max(2, dpr * 2));
}
function downscaleCanvasIfNeeded(srcCanvas, maxW, maxH){
  try{
    const w = srcCanvas.width, h = srcCanvas.height;
    const ratio = Math.min(1, (maxW||w)/w, (maxH||h)/h);
    if(ratio >= 1) return srcCanvas;

    const c = document.createElement("canvas");
    c.width = Math.max(1, Math.round(w * ratio));
    c.height = Math.max(1, Math.round(h * ratio));
    const ctx = c.getContext("2d", { alpha:false });
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(srcCanvas, 0, 0, c.width, c.height);
    return c;
  }catch(e){
    return srcCanvas;
  }
}


/* =========================
   RESET
   ========================= */
function resetAll(){
  if(!confirm("Esto borrará toda la información guardada en este navegador. ¿Continuar?")) return;
  localStorage.removeItem(LS_KEY);
  db = defaultDB();
  saveDB();
  renderStudents(); renderAreas(); fillAreaSelects(); fillStudentSelect();
  renderGradebook(); renderFinalArea(); renderGenerator();
}


/* =========================
   CERTIFICADOS (toma datos del Generador Final)
   ========================= */
function fillCertStudentSelect(){
  const sel = document.getElementById("certStudent");
  if(!sel) return;
  const cur = sel.value;
  sel.innerHTML = db.students.map(s=>`<option value="${s.id}">${escapeHTML(s.name||"")}${s.ident?(" — "+escapeHTML(s.ident)):""}</option>`).join("");
  if(cur && db.students.some(s=>s.id===cur)) sel.value = cur;
  if(!sel.value && db.students[0]) sel.value = db.students[0].id;
}

function fillCertResolution(){
  const inp = document.getElementById("certResolution");
  if(!inp) return;
  inp.value = db.settings.certResolution || "";
}
function saveCertResolution(){
  db.settings.certResolution = (document.getElementById("certResolution")?.value || "").trim();
  saveDB();
  renderCertificate();
}

function computeCertRows(studentId){
  const rows = [];
  db.areas.forEach(a=>{
    const fin = computeAreaFinalForStudent(a.id, studentId);
    const anual = avgSimple(fin.per.map(x=>x.nota));
    rows.push({
      area: a.name||"",
      ihs: a.ihs||"",
      nota: anual,
      des: isNum(anual) ? escalaNacional_SIN_D(anual) : " "
    });
  });
  return rows;
}

function ddmmyyyyToLongES(ddmmyyyy){
  const m = String(ddmmyyyy||"").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  const months = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  if(!m) return String(ddmmyyyy||"").trim();
  const d = Number(m[1]); const mo = Number(m[2]); const y = Number(m[3]);
  const mes = months[clamp(mo,1,12)-1];
  return `a los ${d} días del mes de ${mes} del año ${y}`;
}

function inferMunicipioDepto(place){
  const p = String(place||"").trim();
  if(!p) return { municipio:"", depto:"" };
  const parts = p.split(",").map(x=>x.trim()).filter(Boolean);
  if(parts.length>=2) return { municipio: parts[0], depto: parts.slice(1).join(", ") };
  return { municipio:p, depto:"" };
}


function renderCertBarcode(value){
  const clean = String(value||"").trim();
  const svg = document.getElementById("certBarcode");
  const txt = document.getElementById("certBarcodeText");
  if(txt) txt.textContent = clean;
  if(!svg) return;

  if(window.JsBarcode){
    try{
      JsBarcode(svg, clean || "—", {
        format: "CODE128",
        displayValue: false,
        lineColor: "#000",
        width: 1.25,
        height: 40,
        margin: 0
      });
      return;
    }catch(e){}
  }
  // Fallback si no cargó la librería: muestra solo el número
  svg.outerHTML = `<div class="mono" style="font-size:10px;text-align:right">${escapeHTML(clean||"")}</div>`;
}

function renderCertificate(){
  const wrap = document.getElementById("certWrap");
  if(!wrap) return;

  const studentId = document.getElementById("certStudent")?.value;
  if(!studentId){
    wrap.innerHTML = '<div class="note">No hay estudiante seleccionado.</div>';
    return;
  }
  const st = db.students.find(s=>s.id===studentId);
  if(!st){
    wrap.innerHTML = '<div class="note">Estudiante inválido.</div>';
    return;
  }

  const ie = (db.settings.ieName || db.settings.institution || "").trim();
  const nit = (db.settings.nit || "").trim();
  const dane = (db.settings.dane || "").trim();
  const place = (db.settings.place || "").trim();
  const muniSet = (db.settings.municipio || "").trim();
  const deptoSet = (db.settings.depto || "").trim();
  const infMD = inferMunicipioDepto(place);
  const municipio = muniSet || infMD.municipio;
  const depto = deptoSet || infMD.depto;

  const year = (db.settings.year || "").trim();
  const grado = (st.grade || db.settings.gradeGlobal || "").trim();
  const now = new Date();
  const months = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  const fechaTxt = `a los ${now.getDate()} días del mes de ${months[now.getMonth()]} del año ${now.getFullYear()}`;
  const horaTxt = now.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });
  const rows = computeCertRows(studentId);
  const promFinal = avgSimple(rows.map(r=>r.nota).filter(isNum));
  const desFinal = isNum(promFinal) ? escalaNacional_SIN_D(promFinal) : " ";

const promovido = getBoletinPromovidoForStudent(st);
let alGrado = getBoletinAlGradoForStudent(st);

if(promovido==="SI" && !alGrado){
  // Si el grado es numérico, intenta sugerir el siguiente
  const n = (grado.match(/\d+/)?.[0]) ? Number(grado.match(/\d+/)[0]) : null;
  if(Number.isFinite(n)) alGrado = String(n+1);
}

  const shield = db.settings.images?.shield || "";
  const resAprob = (db.settings.certResolution || "").trim();

  const rector = (db.settings.rectorName || "").trim();
  const dir = (db.settings.dirName || "").trim();

  const sigR = db.settings.images?.rectorSign || "";
  const sigD = db.settings.images?.dirSign || "";

  wrap.innerHTML = `
    <div class="certificado" id="certPaper">
      <div class="c-head">
        <div class="c-barcode">
          <div>
            <svg id="certBarcode"></svg>
            <div class="c-bctext mono" id="certBarcodeText"></div>
          </div>
        </div>
        <div class="c-headgrid">
          <div class="c-escudo">${shield ? `<img src="${shield}" alt="Escudo IE">` : ``}</div>
          <div class="c-headtext">
            <div class="c-l1">REPUBLICA DE COLOMBIA</div>
            <div class="c-l2">${depto ? ("DEPARTAMENTO DE " + escapeHTML(depto).toUpperCase()) : "DEPARTAMENTO"}</div>
            <div class="c-l3">${municipio ? ("MUNICIPIO DE " + escapeHTML(municipio).toUpperCase()) : "MUNICIPIO"}</div>
            <div class="c-inst">${escapeHTML(ie || "INSTITUCIÓN EDUCATIVA")}</div>
            ${nit ? `<div class="c-sub">NIT ${escapeHTML(nit)}</div>` : ``}
            ${dane ? `<div class="c-sub">CÓDIGO DANE ${escapeHTML(dane)}</div>` : ``}
            ${place ? `<div class="c-sub">${escapeHTML(place)}</div>` : ``}</div>
        </div>
      </div>

      <div class="c-text">
        El Rector y el Director de Grado <span class="c-strong">${escapeHTML(grado || "—")}</span> certifican que el(la) estudiante
        <span class="c-strong">${escapeHTML(st.name||"")}</span>, identificado(a) con
        <span class="c-strong">${escapeHTML(st.ident||"—")}</span>, cursó en este establecimiento las áreas obligatorias y optativas establecidas
        en el plan de estudios, correspondientes al grado <span class="c-strong">${escapeHTML(grado || "—")}</span> durante el año lectivo
        <span class="c-strong">${escapeHTML(year || "—")}</span>, de acuerdo con lo establecido en la Ley 115 de 1994, el Decreto 1860 de 1994
        y el Decreto 1290 de 2009.
      </div>

      <div class="c-title">CERTIFICA QUE</div>

      <table class="c-table">
        <thead>
          <tr>
            <th>ÁREA Y/O ASIGNATURA</th>
            <th style="width:70px">IHS</th>
            <th style="width:120px">ESCALA DE VALORACIÓN</th>
            <th style="width:160px">NIVEL DE DESEMPEÑO</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r=>`
            <tr>
              <td>${escapeHTML(r.area)}</td>
              <td class="b-center mono">${escapeHTML(String(r.ihs||""))}</td>
              <td class="b-center mono">${isNum(r.nota)? escapeHTML(fmt(r.nota,2)) : ""}</td>
              <td class="b-center">${escapeHTML(r.des||" ")}</td>
            </tr>
          `).join("")}
        </tbody>
        <tfoot>
          <tr>
            <th colspan="2" class="b-right">PROMEDIO FINAL</th>
            <th class="b-center mono">${isNum(promFinal)? escapeHTML(fmt(promFinal,2)) : ""}</th>
            <th class="b-center">${escapeHTML(desFinal||" ")}</th>
          </tr>
        </tfoot>
      </table>

      <div class="c-prom">
        <div class="box"><b>PROMOVIDO:</b> ${promovido==="SI" ? "SI" : (promovido==="NO" ? "NO" : "ND")}</div>
        ${promovido==="NO" ? `<div class="box"><b>AL GRADO:</b> ${escapeHTML(alGrado||"—")}</div>` : (promovido==="SI" && alGrado ? `<div class="box"><b>AL GRADO:</b> ${escapeHTML(alGrado)}</div>` : ``)}
      </div>

      <div class="c-legend">
        <h3>DESCRIPCIONES GENERALES DE DESEMPEÑO (Decreto 1290 de 2009)</h3>
        <div class="lrow"><b>DESEMPEÑO BAJO (1,00 - 2,99)</b><span>No supera los desempeños necesarios en relación con los objetivos y metas establecidas.</span></div>
        <div class="lrow"><b>DESEMPEÑO BÁSICO (3,00 - 3,99)</b><span>Logra lo mínimo en relación con los niveles de aprendizaje esperados.</span></div>
        <div class="lrow"><b>DESEMPEÑO ALTO (4,00 - 4,59)</b><span>Alcanza la mayoría de los desempeños, evidenciando un buen nivel de desarrollo.</span></div>
        <div class="lrow"><b>DESEMPEÑO SUPERIOR (4,60 - 5,00)</b><span>Alcanza la totalidad de los desempeños, superando los objetivos y metas.</span></div>
      </div>

      <div class="c-foot">
        Se expide sin enmendaduras de ninguna índole en Balboa Cauca ${fechaTxt ? (", " + escapeHTML(fechaTxt) + (horaTxt ? (" a las " + escapeHTML(horaTxt)) : "")) : ""}.
      </div>

      <div class="c-signgrid">
        <div class="c-sign">
          <div class="sigimg">${sigR ? `<img src="${sigR}" alt="Firma Rector">` : ``}</div>
          <div class="line">
            <div class="name">${escapeHTML(rector || "—")}</div>
            <div class="role">Rector</div>
          </div>
        </div>
        <div class="c-sign">
          <div class="sigimg">${sigD ? `<img src="${sigD}" alt="Firma Director de Grupo">` : ``}</div>
          <div class="line">
            <div class="name">${escapeHTML(dir || "—")}</div>
            <div class="role">Director(a) de grado</div>
          </div>
        </div>
      </div>
    </div>
  `;
  renderCertBarcode(st.ident);
}

function setPrintMode(mode){
  document.body.dataset.print = mode;
}
function clearPrintMode(){
  delete document.body.dataset.print;
}
function printBoletin(){
  const prevTitle = document.title;
  try{
    const sid = document.getElementById("genStudent")?.value;
    const st = db.students.find(s=>s.id===sid);
    const name = safeFileName(st?.name || "ESTUDIANTE").toUpperCase();
    const per = Number(document.getElementById("genPeriod")?.value || 1) || 1;
    const gradeTok = normalizeGradeToken(st?.grade || db.settings.gradeGlobal || "");
    document.title = `${name}_BOLETIN_PER${per}_GRADO${gradeTok}`;
  }catch(e){}
  setPrintMode("boletin");
  window.onafterprint = ()=>{ document.title = prevTitle; clearPrintMode(); };
  window.print();
}
function printCertificado(){
  const prevTitle = document.title;
  try{
    const sid = document.getElementById("certStudent")?.value;
    const st = db.students.find(s=>s.id===sid);
    const name = safeFileName(st?.name || "ESTUDIANTE").toUpperCase();
    const gradeTok = normalizeGradeToken(st?.grade || db.settings.gradeGlobal || "");
    document.title = `${name}_CERTIFICADO_GRADO${gradeTok}`;
  }catch(e){}
  setPrintMode("cert");
  window.onafterprint = ()=>{ document.title = prevTitle; clearPrintMode(); };
  window.print();
}


async function exportBoletinPDFGeneradorFinal(){
  const studentId = document.getElementById("genStudent")?.value;
  if(!studentId) return alert("Selecciona un estudiante.");
  if(!ensurePDFLibs()) return alert("No se cargaron librerías de PDF (html2canvas/jspdf).");

  const paper = document.querySelector("#genTableWrap .boletin");
  if(!paper) return alert("No hay boletín para exportar.");

  try{
    // Render con alta nitidez pero tamaño de archivo controlado
    const dpr = window.devicePixelRatio || 1;
    const scale = getPDFRenderScale(dpr);
    let canvas = await html2canvas(paper, {
      scale,
      useCORS: true,
      backgroundColor: "#ffffff",
      logging: false
    });
    // Limitar a aprox. 300 DPI en carta (reduce peso sin perder legibilidad)
    canvas = downscaleCanvasIfNeeded(canvas, 2550, 3300);
    const imgData = canvas.toDataURL("image/jpeg", 0.92);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation:"p", unit:"pt", format:"letter", compress:true });

    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    const imgW = canvas.width;
    const imgH = canvas.height;

    const ratio = Math.min(pageW / imgW, pageH / imgH);
    const w = imgW * ratio;
    const h = imgH * ratio;
    const x = (pageW - w) / 2;
    const y = 24;

    pdf.addImage(imgData, "JPEG", x, y, w, h, undefined, "FAST");

    const st = db.students.find(s=>s.id===studentId);
    const name = safeFileName(st?.name || "ESTUDIANTE").toUpperCase();
    const gradeTok = normalizeGradeToken(st?.grade || db.settings.gradeGlobal || "");
    const mode = document.getElementById("genMode")?.value || "final";
    const perTok = (mode === "final")
      ? "FINAL"
      : safeFileName(document.getElementById("genPeriod")?.value || "PERIODO").toUpperCase();

    pdf.save(`${name}_${gradeTok}_${perTok}.pdf`);
  }catch(e){
    console.error(e);
    alert("No se pudo exportar el boletín a PDF.");
  }
}

async function exportCertificadoPDF(){
  const studentId = document.getElementById("certStudent")?.value;
  if(!studentId) return alert("Selecciona un estudiante.");
  if(!ensurePDFLibs()) return alert("No se cargaron librerías de PDF (html2canvas/jspdf).");

  const paper = document.getElementById("certPaper");
  if(!paper) return alert("No hay certificado para exportar.");

  try{
    // Render con alta nitidez pero tamaño de archivo controlado
    const dpr = window.devicePixelRatio || 1;
    const scale = getPDFRenderScale(dpr);
    let canvas = await html2canvas(paper, {
      scale,
      useCORS: true,
      backgroundColor: "#ffffff",
      logging: false
    });
    // Limitar a aprox. 300 DPI en carta (reduce peso sin perder legibilidad)
    canvas = downscaleCanvasIfNeeded(canvas, 2550, 3300);
    const imgData = canvas.toDataURL("image/jpeg", 0.92);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation:"p", unit:"pt", format:"letter", compress:true });

    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    const imgW = canvas.width;
    const imgH = canvas.height;

    const ratio = Math.min(pageW / imgW, pageH / imgH);
    const w = imgW * ratio;
    const h = imgH * ratio;
    const x = (pageW - w) / 2;
    const y = 24;

    pdf.addImage(imgData, "JPEG", x, y, w, h, undefined, "FAST");

    const st = db.students.find(s=>s.id===studentId);
    const name = safeFileName(st?.name || "CERTIFICADO").toUpperCase();
    const gradeTok = normalizeGradeToken(st?.grade || db.settings.gradeGlobal || "");
    pdf.save(`${name}_CERTIFICADO_GRADO${gradeTok}.pdf`);
  }catch(e){
    console.error(e);
    alert("No se pudo exportar el certificado a PDF.");
  }
}

/* =========================
   INIT
   ========================= */
(function init(){
  refreshAllMeta();
  renderStudents();
  renderAreas();
  fillAreaSelects();
  fillPeriodSelects();
  fillStudentSelect();
  fillPeriodSelectsGenerator();
  renderGradebook();
  renderFinalArea();
  renderGenerator();
  go("dashboard");
})();


/* =========================
   IE / LS_KEY: Inicialización UI
   ========================= */
window.addEventListener("load", ()=>{
  try{ updateHeaderIEBadge(); }catch(e){}
  // Primera vez: pedir el LS_KEY para que no se mezclen datos entre IE
  try{
    const hasActive = !!sanitizeLSKey(localStorage.getItem(ACTIVE_KEY_STORE));
    if(!hasActive) openIESwitcher();
  }catch(e){}
});

