/**
 * ============================================================================
 *  SEGUIMIENTO VENTAS Y PERFORMANCE AEROPUERTO  ·  Google Apps Script Web App
 * ============================================================================
 *  Lee la Base de Comisiones + Turnos 360 (hojas "Coordinadores Mayo",
 *  "Agentes Junio" [agentes = ejecutivos de venta], "Anfitriones ...",
 *  "Supervisores ..."). Cada celda de turno es "05:00 - 16:00" / "21:00 - 08:00"
 *  / "Libre".
 *
 *  Reglas:
 *   - ROL del vendedor: se deduce de en que hoja de Turnos aparece (o del Roster).
 *     Secciones separadas: EJECUTIVOS (sus ventas) y COORDINADORES (comision).
 *   - COMISION coordinadores: por sus HORAS ACTIVAS (turno menos horas
 *     administrativas/colacion del Cronograma V77). 2 activos -> 50/50 (N -> 1/N).
 *   - EQUIPO Diurno/Nocturno = turno del EJECUTIVO vendedor (cruza medianoche =
 *     Nocturno). NO depende de la franja: una misma hora puede tener ventas
 *     diurnas y nocturnas a la vez (un ejecutivo diurno y uno nocturno).
 *   - "Otras Ventas": ventas hechas por NO-ejecutivos (supervisor/coordinador/
 *     anfitrion/otro).
 *   - Correos mal ingresados: atributo ds_user_email (solo van compartida).
 *   - VENTAS ROUND (ida y vuelta): compartida = las dos piernas comparten el
 *     reservation_id y la vuelta trae sufijo "Z"; exclusiva = mismo agente,
 *     mismo precio, <=5 min entre creaciones y tm_start con >24 h de diferencia.
 *     Un par = UNA venta round, contabilizada en la pierna creada primero.
 *   - METAS Y TRAMOS POR PERIODO: se guardan por mes (YYYY-MM) con vigencia
 *     hacia adelante. Un periodo sin configuracion propia hereda la del ultimo
 *     periodo anterior configurado, asi cambiar agosto no altera julio.
 *   - ACCESO: los agentes de venta solo ven su propia Vista Agente; los
 *     coordinadores solo ven su propia Vista Coordinador; el resto ve todo.
 * ============================================================================
 */

var CONFIG = {
  VENTAS_SPREADSHEET_ID: '1i5TjTE34M8jeGYKJ7jBbw8KqS6k5YpOTMr5mEO-jR00',
  VENTAS_SHEET: 'Results',

  TURNOS_SPREADSHEET_ID: '1xepmv4-ocTNZ-RXBa7pFBXp4KM56ZTrAgeO5M0B7yCg',
  ANIO: 2026,
  PREFIJO_AGENTES: 'Agentes',           // = ejecutivos de venta
  PREFIJO_COORDINADORES: 'Coordinadores',
  PREFIJO_ANFITRIONES: 'Anfitriones',
  PREFIJO_SUPERVISORES: 'Supervisores',

  ROSTER_SHEET: 'Roster',
  // El Roster es OPCIONAL: si la pestana no existe, la app resuelve todo desde
  // Turnos 360 (rol por mes, nombre y equipo) cruzando por el correo. El Roster,
  // si existe, actua como refuerzo para correos que no sigan ninguna convencion.
  // Alias manual (opcional) correo -> nombre tal como aparece en Turnos 360:
  ALIAS_CORREOS: { /* 'correo.raro@cabify.com': 'Nombre En Turnos' */ },
  // Libro donde esta la pestana Roster. Si la dejas '', usa el libro donde esta
  // instalado este script (o, si no, el de ventas). Si tu Roster esta en OTRO
  // libro (p.ej. "[CL] AIRPORT SALES"), pega aqui su ID (lo sacas de la URL).
  ROSTER_SPREADSHEET_ID: '',

  // (Opcional) Listas blancas de correos para el control de acceso. Si un correo
  // aparece aqui, se fuerza ese rol de acceso aunque el match por nombre falle.
  // Son un refuerzo: dejalas vacias para depender solo de Turnos/Maestros/Roster.
  ACCESO_AGENTES: [ /* 'valery.becerra@cabify.com', ... */ ],
  ACCESO_COORDINADORES: [ /* 'coordinador@cabify.com', ... */ ],

  COL_FECHA_VENTA: 'createdAt_local',      // fecha de la VENTA = creacion de la reserva (coincide con Looker)
  COL_FECHA_FALLBACK: 'tm_start_local_at', // respaldo si faltara createdAt
  PRODUCTO_COMPARTIDA: 'van_compartida',
  PRODUCTO_EXCLUSIVA:  'van_exclusive',
  FINISH_OK: 'FINISH_REASON_DROPOFF',
  // Para comisiones se cuenta BRUTO: viajes que concretan = DROPOFF + ADMIN_CANCEL.
  FINISH_CONCRETADAS: ['FINISH_REASON_DROPOFF','FINISH_REASON_ADMIN_CANCEL'],
  SOLO_CONCRETADAS: true,

  // --- Deteccion de ventas ida y vuelta (Round) ---
  ROUND: {
    SUFIJO: 'Z',            // compartida: la vuelta = mismo reservation_id + este sufijo
    MIN_CREATED: 5,         // minutos maximos entre las dos creaciones
    EXCL_MIN_START_H: 24    // exclusivas: horas minimas de diferencia entre los tm_start
  },

  // Cronograma V77: horas administrativas/colacion por hora de ingreso (no comisionan).
  ADMIN_POR_INGRESO: { 5:[11,12,13], 10:[10,14,15], 13:[17,18,19], 21:[6,7,8] },

  META_PROP: 'METAS_VENTAS_JSON',              // (legado) meta unica, se migra a periodos
  TRAMOS_PROP: 'TRAMOS_COMISION_JSON',         // (legado) tramos unicos, se migran a periodos
  META_PERIODOS_PROP: 'METAS_PERIODOS_JSON',   // { "2026-08": {grupo:...}, "default": {...} }
  TRAMOS_PERIODOS_PROP: 'TRAMOS_PERIODOS_JSON',// { "2026-08": {coord:[],ejec:[]}, "default": {...} }

  // --- Desempeno de agentes (Tableau crosstab) ---
  TABLEAU_SERVER:      'https://tableau.cabify-data.com',
  TABLEAU_API_VERSION: '3.19',
  TABLEAU_SITE:        'cabify',
  TABLEAU_PAT_NAME:    'B2B Support',
  TABLEAU_PAT_SECRET:  'KI/Fpf9RTVaA8F2oH/vqKQ==:kL07mNjXKGkJcnPJT4v4BjB0jXxmOFuI',
  // contentUrl de la vista: workbook/sheets/vista (de tu link PerformanceSupportCrosstab / Sheet1)
  TABLEAU_VIEW_CONTENT_URL: 'PerformanceSupportCrosstab/sheets/Sheet1',
  PERF_SHEET: 'Performance',            // pestana cache donde se vuelca la crosstab
  PERF_PROP:  'PERF_ULTIMA_ACTUALIZACION',
  FIRT_UMBRAL_H: 24,    // cumple si la primera respuesta es antes de 24 horas
  FURT_UMBRAL_H: 120,   // cumple si la resolucion es antes de 120 horas

  // --- Auditorias de agentes (Tableau: Support_CL_Audits / Sheet1) ---
  TABLEAU_AUDIT_CONTENT_URL: 'Support_CL_Audits/sheets/Sheet1',
  AUDIT_SHEET: 'Auditorias',            // pestana cache de auditorias (Email, Ticket, Fecha, Score)
  AUDIT_PROP:  'AUDIT_ULTIMA_ACTUALIZACION',
  AUDIT_MAX_AGE_H: 12,                  // si la cache de auditorias es mas vieja, se re-sincroniza sola
  // La nota de auditoria NO se ata al periodo de ventas: 'Created At Dttm UTC' es la fecha del
  // TICKET (QA audita con semanas de desfase), por eso un periodo corto deja casi sin datos.
  // 0 = promedia TODAS las auditorias en cache; >0 = solo las de los ultimos N dias (por fecha de ticket).
  AUDIT_VENTANA_DIAS: 0,
  AUDIT_PDF_FOLDER: '',                 // (opcional) ID de carpeta de Drive para los PDF; '' = raiz

  // --- Datos maestros (planilla BUK de colaboradores activos) ---
  MAESTROS_SHEET: 'Maestros',           // pestana cache sincronizada desde la planilla
  MAESTROS_PROP:  'MAESTROS_SYNC',

  // --- Ausentismo / incidencias (libro que alimenta el Apps Script "BSC Aeropuerto") ---
  BSC_SPREADSHEET_ID: '16AWUUvKZOeziyQx3Qxsb76U28zUPAA6jtbisruPGyw4',
  BSC_CONSOLIDADO_SHEET: 'Consolidado',
  BSC_COLAB_SHEET: 'Colaboradores',
  BSC_ALIAS_SHEET: 'Alias',
  MANUALES_SHEET: 'IncidenciasManuales',   // incidencias registradas manualmente por supervisores (mandan sobre BSC)

  // --- Modelo de comision de calidad (base = ventas x %tramo; 70% fijo + 30% variable en 3 criterios de 10%) ---
  // Por criterio: si Real < inflexion => 0%; si no => min(Real/Meta, 1). Bono = %cumpl x 10% x base.
  CALIDAD: {
    registros:  { meta:0.90, inflexion:0.80 },   // % de registros (correos) correctos
    auditoria:  { meta:0.90, inflexion:0.70 },   // nota auditoria /100 (sin auditorias => 100%)
    adherencia: { meta:0.90, inflexion:0.80 }    // turnos sin incidencia que descuenta / turnos trabajables (sin datos => 100%)
  },
  SUPERVISORES_PROP: 'SUPERVISORES_JSON',        // lista de correos que pueden auditar/sobrescribir incidencias
  MOSTRAR_COMISION_PROP: 'MOSTRAR_COMISION_JSON', // visibilidad de la comision en la vista de agente/coordinador
  OVERRIDES_SHEET: 'OverridesIncidencias'        // registro de sobrescrituras de incidencias
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate().setTitle('Seguimiento Ventas y Performance · Aeropuerto')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ------------------------------- CACHE DE LECTURAS -------------------------
// Cachea (5 min, a nivel de script) las lecturas pesadas (ventas, turnos, roles)
// para que la PRIMERA carga las guarde y las otras vistas (agente/coordinador)
// las reutilicen sin volver a leer las hojas. Se invalida al sincronizar maestros
// o al pulsar "Actualizar" en el dashboard.
function _cacheVer_(){ try{ return PropertiesService.getScriptProperties().getProperty('CACHE_VER')||'1'; }catch(e){ return '1'; } }
function resetCacheDatos(){ try{ var p=PropertiesService.getScriptProperties();
  var v=(parseInt(p.getProperty('CACHE_VER')||'1',10)+1); p.setProperty('CACHE_VER', String(v)); return v; }catch(e){ return 0; } }
function _cacheGet_(key){
  try{ var c=CacheService.getScriptCache(), k=key+'@'+_cacheVer_();
    var meta=c.get(k+'_n'); if(!meta) return null;
    var n=+meta, keys=[]; for(var i=0;i<n;i++) keys.push(k+'_'+i);
    var parts=c.getAll(keys), s='';
    for(var j=0;j<n;j++){ var p=parts[k+'_'+j]; if(p==null) return null; s+=p; }
    return JSON.parse(s);
  }catch(e){ return null; }
}
function _cachePut_(key, obj, ttl){
  try{ var c=CacheService.getScriptCache(), k=key+'@'+_cacheVer_();
    var s=JSON.stringify(obj), CH=95000, n=Math.ceil(s.length/CH);
    if(n>95) return;   // demasiado grande para cachear -> se lee cada vez
    var o={}; for(var i=0;i<n;i++) o[k+'_'+i]=s.substring(i*CH,(i+1)*CH);
    o[k+'_n']=String(n); c.putAll(o, ttl||300);
  }catch(e){}
}

// ============================== CONTROL DE ACCESO ==========================

/** Traduce errores tecnicos a mensajes claros para el usuario final. */
function _amable_(e){
  var m=String((e&&e.message)||e||'');
  if(m.indexOf('You do not have permission')>=0 || m.indexOf('permission')>=0 || m.indexOf('PERMISSION')>=0)
    return 'Tu cuenta no tiene permisos para leer los datos de la app. Esto se corrige en la publicacion de la Web App: debe estar configurada como "Ejecutar como: propietario". Avisa al administrador.';
  if(m.indexOf('Service invoked too many times')>=0 || m.indexOf('quota')>=0)
    return 'La app alcanzo el limite temporal de consultas de Google. Espera un par de minutos y vuelve a intentar.';
  if(m.indexOf('Timeout')>=0 || m.indexOf('timed out')>=0)
    return 'La carga tardo demasiado. Intenta nuevamente en unos segundos.';
  return m.replace(/^Exception:\s*/,'');
}

// Bandera de ejecucion interna: getMesAgentes/getMesCoordinadores llaman a
// getDashboard por dentro (necesitan sus datos) aunque quien invoca sea un
// usuario restringido. La bandera es estado del servidor dentro de UNA
// invocacion, no un parametro que el cliente pueda enviar, por eso no es
// evadible desde google.script.run.
var _INTERNAL_BYPASS_ = false;
var _ACC_CACHE_ = null;   // cache del contexto de acceso por invocacion

/** Identidad y rol de acceso del usuario que abrio la app.
 *  Devuelve {email, esAgente, esCoordinador, nombre, nk, identificado}.
 *  esAgente / esCoordinador => acceso restringido a su propia vista. */
function getAccessContext(){
  if(_ACC_CACHE_) return _ACC_CACHE_;
  var email='';
  try { email=(Session.getActiveUser().getEmail()||'').trim().toLowerCase(); } catch(e){}
  if(!email){
    // Sin identidad (deploy anonimo): no se puede restringir -> acceso completo.
    _ACC_CACHE_={email:'', esAgente:false, esCoordinador:false, esSupervisor:false, nombre:'', nk:'', identificado:false};
    return _ACC_CACHE_;
  }
  try {
    var ym=dateKey_(new Date()).substring(0,7);
    var RB=buildRoles_(), roleKM=RB.roleKM, nameByKey=RB.nameByKey;
    var maestros=loadMaestros_();
    var R=loadRoster_();

    // candidatos de nameKey a partir del correo (mismo criterio que infoVendedor)
    var t0=norm_(String(email).split('@')[0].replace(/_/g,'.').split('.').filter(Boolean)[0]||'');
    var concatIdx={}, firstIdx={};
    Object.keys(nameByKey).forEach(function(k){ var p=k.split(' ');
      (firstIdx[p[0]]=firstIdx[p[0]]||[]).push(k); if(p.length>=2) concatIdx[p[0]+p[p.length-1]]=k; });
    var toks=norm_(String(email).split('@')[0]).split(/[._]/).filter(Boolean);
    var mm=null; if(toks.length){ var hits=maestros.filter(function(m){
        return toks.every(function(t){ return m.tokSet[t]||m.pairSet[t]; }); }); if(hits.length===1) mm=hits[0]; }

    var cands=[];
    if(CONFIG.ALIAS_CORREOS[email]) cands.push(nameKey_(CONFIG.ALIAS_CORREOS[email]));
    if(R.keyByEmail[email]) cands.push(R.keyByEmail[email]);
    if(mm) cands.push(mm.key);
    cands.push(keyFromEmail_(email));
    if(concatIdx[t0]) cands.push(concatIdx[t0]);
    if(firstIdx[t0] && firstIdx[t0].length===1) cands.push(firstIdx[t0][0]);
    var seen={}; cands=cands.filter(function(k){ if(!k||seen[k])return false; seen[k]=1; return true; });

    var nk=null; for(var i=0;i<cands.length;i++){ if(nameByKey[cands[i]]){ nk=cands[i]; break; } }
    if(!nk) nk=cands[0];
    var rol=''; for(var j=0;j<cands.length;j++){ if(roleKM[cands[j]+'|'+ym]){ rol=roleKM[cands[j]+'|'+ym]; break; } }
    if(!rol) rol=(mm&&mm.rol)||(R.byEmail[email]&&R.byEmail[email].rol)||'Otro';
    var nombre=(mm&&mm.nombre)||(R.byEmail[email]&&R.byEmail[email].nombre)||nameByKey[nk]||email;

    // Listas blancas de refuerzo (por si el match por nombre no cubre a alguien)
    var wlAg=(CONFIG.ACCESO_AGENTES||[]).map(function(x){return String(x).trim().toLowerCase();});
    var wlCo=(CONFIG.ACCESO_COORDINADORES||[]).map(function(x){return String(x).trim().toLowerCase();});
    var esSup=(getSupervisores().indexOf(email)>=0);   // supervisor autorizado (lista cargada en config)
    var esAgente=(rol==='Ejecutivo') || wlAg.indexOf(email)>=0;
    var esCoord =(rol==='Coordinador') || wlCo.indexOf(email)>=0;
    if(esCoord) esAgente=false;
    if(esSup){ esAgente=false; esCoord=false; }   // el supervisor NO es un perfil restringido: ve todo y puede auditar

    _ACC_CACHE_={email:email, esAgente:esAgente, esCoordinador:esCoord, esSupervisor:esSup, nombre:nombre, nk:nk, identificado:true};
  } catch(err){
    // La cuenta no pudo LEER los libros (tipico: deploy "como usuario que accede"
    // sin permisos sobre los Sheets). No reventar: devolver contexto con flag de
    // error para que el front muestre un mensaje limpio en vez de la excepcion.
    _ACC_CACHE_={email:email, esAgente:false, esCoordinador:false, esSupervisor:false, nombre:'', nk:'',
      identificado:true, errorPermisos:true, mensaje:_amable_(err)};
  }
  return _ACC_CACHE_;
}

/** Lanza error si el usuario es un perfil restringido (agente o coordinador).
 *  Se usa como candado en las funciones que solo debe ver el resto de usuarios. */
function _bloquearSiRestringido_(){
  var acc=getAccessContext();
  if(acc.esAgente || acc.esCoordinador)
    throw new Error('Acceso restringido: tu usuario solo tiene acceso a tu vista personal.');
}

// ------------------------------- utilidades --------------------------------
function ss_(id){ return SpreadsheetApp.openById(id || CONFIG.VENTAS_SPREADSHEET_ID); }
function norm_(s){ return (s==null?'':String(s)).normalize('NFKD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase(); }
function pad2_(n){ return (n<10?'0':'')+n; }
function dateKey_(d){ return d.getFullYear()+'-'+pad2_(d.getMonth()+1)+'-'+pad2_(d.getDate()); }
function _hhmm_(v){
  if(v==null||v==='') return '';
  if(v instanceof Date) return pad2_(v.getHours())+':'+pad2_(v.getMinutes());
  var s=String(v).trim(); if(!s||s==='-'||s==='None') return '';
  var num=parseFloat(s);
  if(!isNaN(num) && num>0 && num<1){ var tot=Math.round(num*86400); return pad2_(Math.floor(tot/3600))+':'+pad2_(Math.floor((tot%3600)/60)); }
  var m=s.match(/(\d{1,2}):(\d{2})/); if(m) return pad2_(+m[1])+':'+pad2_(+m[2]);
  return s;
}
function addDays_(d,n){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()+n); }
function round_(n,d){ d=(d==null?0:d); var f=Math.pow(10,d); return Math.round((Number(n)||0)*f)/f; }
function firstOfMonth_(){ var n=new Date(); return new Date(n.getFullYear(), n.getMonth(), 1); }
function endOfDay_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23,59,59); }
function getColMap_(h){ var m={}; h.forEach(function(x,i){ m[String(x).trim()]=i; }); return m; }
function nameKey_(s){ var t=norm_(s).split(/\s+/).filter(Boolean); return !t.length?'':(t.length===1?t[0]:t[0]+' '+t[t.length-1]); }
function keyFromEmail_(e){ var t=String(e).split('@')[0].split(/[._]/).filter(Boolean); return !t.length?'':(t.length===1?norm_(t[0]):norm_(t[0])+' '+norm_(t[t.length-1])); }

function parseDate_(v, mdy){
  if (v instanceof Date && !isNaN(v)) return v;
  if (v==null || v==='') return null;
  var s=String(v).trim();
  // ISO: yyyy-mm-dd [hh:mm[:ss]] -> se interpreta como hora LOCAL (no UTC), para
  // que el selector de fechas y los datos no se corran un dia en zonas UTC-.
  var iso=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3], iso[4]?+iso[4]:0, iso[5]?+iso[5]:0, iso[6]?+iso[6]:0);
  // d/m/yyyy (o d-m-yyyy). Con mdy=true se interpreta MES primero (formato US de Tableau).
  var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m){ var dd=mdy?+m[2]:+m[1], mm=mdy?+m[1]:+m[2];
    return new Date(+m[3], mm-1, dd, m[4]?+m[4]:0, m[5]?+m[5]:0); }
  m=s.toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'').match(/^(\d{1,2}) de ([a-z]+) de (\d{4})$/);
  if (m && MES_FULL_[m[2]]!=null) return new Date(+m[3], MES_FULL_[m[2]], +m[1]);
  var d=new Date(s); return isNaN(d)?null:d;
}
function parseTurno_(v){
  if (v==null) return {libre:true};
  var s=String(v).trim().toLowerCase();
  if (!s||s==='libre'||s==='.'||s==='-'||s.indexOf('vacac')>=0) return {libre:true};
  var m=s.match(/(\d{1,2})[:\.]?\d{0,2}\s*[-a]\s*(\d{1,2})[:\.]?\d{0,2}/);
  if (!m) return {libre:true};
  var st=+m[1], en=+m[2]; if (st>23||en>24) return {libre:true};
  return { libre:false, start:st, end:en, cross:(en<=st) };
}
var MES_={ene:0,feb:1,mar:2,abr:3,may:4,jun:5,jul:6,ago:7,sep:8,set:8,oct:9,nov:10,dic:11};
function parseHeaderDate_(v){
  if (v instanceof Date && !isNaN(v)) return v;
  if (v==null) return null;
  var s=String(v).trim().toLowerCase();
  var m=s.match(/^(\d{1,2})[\-\/ ]([a-záéíóú]{3,})/);
  if (m && MES_[m[2].substring(0,3)]!=null) return new Date(CONFIG.ANIO, MES_[m[2].substring(0,3)], +m[1]);
  m=s.match(/^(\d{1,2})[\-\/](\d{1,2})$/);
  if (m) return new Date(CONFIG.ANIO, (+m[2])-1, +m[1]);
  return null;
}

// ------------------------------- validez email -----------------------------
var VALID_DOM_={'gmail.com':1,'hotmail.com':1,'hotmail.es':1,'hotmail.cl':1,'outlook.com':1,'outlook.es':1,'outlook.cl':1,'yahoo.com':1,'yahoo.es':1,'yahoo.cl':1,'yahoo.com.ar':1,'yahoo.com.br':1,'icloud.com':1,'live.com':1,'live.cl':1,'me.com':1,'mail.com':1,'msn.com':1,'vtr.net':1,'uc.cl':1,'usach.cl':1,'udd.cl':1,'miuandes.cl':1,'ug.uchile.cl':1,'fen.uchile.cl':1,'mayor.cl':1,'latam.com':1,'duocuc.cl':1,'santotomas.cl':1,'unab.cl':1,'protonmail.com':1,'aol.com':1,'proton.me':1,'flesan.cl':1};
var TYPO_DOM_={'gamil.com':1,'gmial.com':1,'gmai.com':1,'gmail.om':1,'gmailc.om':1,'gmaill.com':1,'gmail.cmo':1,'gail.com':1,'gmil.com':1,'gnail.com':1,'gmal.com':1,'gmail.co':1,'gmail.cl':1,'gmaul.com':1,'gmali.com':1,'hotmial.com':1,'hotmai.com':1,'hotmal.com':1,'outlok.com':1,'yaho.com':1,'notiene.cl':1,'gmail.es':1};
var PLACEHOLDER_={'test':1,'prueba':1,'na':1,'no':1,'notiene':1,'nocorreo':1,'sincorreo':1,'xxx':1,'asdf':1,'qwerty':1,'none':1,'nomail':1,'noemail':1,'sin':1,'correo':1,'x':1,'aaa':1};
var TLD_OK_={'com':1,'cl':1,'es':1,'net':1,'org':1,'ar':1,'br':1,'pe':1,'co':1,'io':1,'edu':1,'gob':1,'gov':1};
function emailOk_(e){
  e=(e==null?'':String(e)).trim().toLowerCase();
  if(!e||e==='nan'||e.indexOf('@')<0) return false;
  if(!/^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/.test(e)) return false;
  var p=e.split('@'),loc=p[0],dom=p[1];
  if(TYPO_DOM_[dom]) return false;
  if(!VALID_DOM_[dom]){ if(/(gm|gn|ga|hotm|outl|yah)/.test(dom)) return false; if(!TLD_OK_[dom.split('.').pop()]) return false; }
  if(/^\d+$/.test(loc)) return false;
  var u={}; for(var i=0;i<loc.length;i++) u[loc[i]]=1; if(Object.keys(u).length===1) return false;
  if(loc.length<=2) return false;
  if(PLACEHOLDER_[loc]) return false;
  if(/^[a-z]+$/.test(loc)&&loc.length<=6&&loc.indexOf('.')<0) return false;
  return true;
}

// ------------------------------- carga -------------------------------------
function rosterBook_(){
  if(CONFIG.ROSTER_SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.ROSTER_SPREADSHEET_ID);
  try{ var a=SpreadsheetApp.getActiveSpreadsheet(); if(a) return a; }catch(e){}
  return ss_(CONFIG.VENTAS_SPREADSHEET_ID);
}
function loadRoster_(){
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.ROSTER_SHEET)||book.getSheetByName('Roster seed')||book.getSheetByName('Roster_seed');
  var byEmail={}, keyByEmail={};
  if(!sh) return {byEmail:byEmail, keyByEmail:keyByEmail};
  var v=sh.getDataRange().getValues(); if(v.length<2) return {byEmail:byEmail, keyByEmail:keyByEmail};
  // Si el CSV quedo pegado en una sola columna (no dividido por comas), dividirlo.
  if(v[0].length<2 || (String(v[0][0]).indexOf(',')>=0 && (''+(v[0][1]||'')).trim()==='')){
    v=v.map(function(r){ return String(r[0]==null?'':r[0]).split(','); });
  }
  // Buscar la fila de encabezado (por si hay filas arriba): la que contenga 'Email_Cabify'.
  var hr=-1; for(var k=0;k<Math.min(8,v.length);k++){
    if(v[k].some(function(x){return String(x).trim()==='Email_Cabify';})){ hr=k; break; } }
  if(hr<0) hr=0;
  var c=getColMap_(v[hr]);
  if(c['Email_Cabify']==null) return {byEmail:byEmail, keyByEmail:keyByEmail};
  for(var i=hr+1;i<v.length;i++){
    var email=String(v[i][c['Email_Cabify']]||'').trim().toLowerCase(); if(!email||email.indexOf('@')<0) continue;
    var nombre=v[i][c['Nombre']]||email;
    byEmail[email]={nombre:nombre, rol:(v[i][c['Rol']]||'').toString().trim(), cargo:v[i][c['Cargo']]||'', supervisor:v[i][c['Supervisor']]||''};
    keyByEmail[email]=nameKey_(nombre);
  }
  return {byEmail:byEmail, keyByEmail:keyByEmail};
}

function loadVentas_(){
  var cached=_cacheGet_('VEN2');
  if(cached) return cached.map(function(r){ var d=new Date(r.t);
    return {date:d, hour:d.getHours(), dateKey:dateKey_(d), agent:r.a, user:r.u, amount:r.m, product:r.p, fin:r.f,
            rid:r.r, start:r.s}; });
  var sh=ss_(CONFIG.VENTAS_SPREADSHEET_ID).getSheetByName(CONFIG.VENTAS_SHEET);
  if(!sh) throw new Error('No existe la pestana de ventas: '+CONFIG.VENTAS_SHEET);
  var v=sh.getDataRange().getValues(), c=getColMap_(v[0]);
  var iD=c[CONFIG.COL_FECHA_VENTA], iC=c[CONFIG.COL_FECHA_FALLBACK], iA=c['ds_agent_email'], iU=c['ds_user_email'],
      iM=c['qt_price_local'], iP=c['ds_product_name'], iF=c['finishReason'],
      iR=c['reservation_id'], iS=c['tm_start_local_at'];
  var rows=[], slim=[];
  for(var i=1;i<v.length;i++){
    var d=parseDate_(v[i][iD]); if(!d && iC!=null) d=parseDate_(v[i][iC]); if(!d) continue;
    var ag=String(v[i][iA]||'').trim().toLowerCase(), us=String(v[i][iU]||'').trim().toLowerCase(),
        am=Number(v[i][iM])||0, pr=String(v[i][iP]||'').trim(), fn=String(v[i][iF]||'').trim(),
        rid=(iR!=null? String(v[i][iR]==null?'':v[i][iR]).trim() : '');
    var st=(iS!=null? parseDate_(v[i][iS]) : null); st = st? st.getTime() : 0;
    rows.push({date:d, hour:d.getHours(), dateKey:dateKey_(d), agent:ag, user:us, amount:am, product:pr, fin:fn,
               rid:rid, start:st});
    slim.push({t:d.getTime(), a:ag, u:us, m:am, p:pr, f:fn, r:rid, s:st});
  }
  _cachePut_('VEN2', slim, 300);
  return rows;
}

// --------------------- VENTAS ROUND (ida y vuelta) --------------------------
/** Marca en cada venta: round (las dos piernas del par) y roundMain (la pierna
 *  creada primero = la que representa LA VENTA del par, para no contar doble).
 *  Solo se emparejan viajes concretados con precio > 0. */
function _concretadaR_(fin){ return CONFIG.FINISH_CONCRETADAS.indexOf(fin)>=0; }
function _ridNorm_(v){ return (v==null?'':String(v)).trim(); }
function _ridBase_(rid){
  var s=_ridNorm_(rid), suf=CONFIG.ROUND.SUFIJO;
  if(s.length>suf.length && s.slice(-suf.length).toUpperCase()===suf.toUpperCase()) return s.slice(0,-suf.length);
  return s;
}
function _marcarPar_(a,b,tipo){
  a.round=true; b.round=true; a.roundTipo=tipo; b.roundTipo=tipo;
  var main=(a.date<=b.date)?a:b; main.roundMain=true;
}

function detectarRoundTrips_(ventas){
  var MS_MIN=60000, MS_H=3600000, CR=CONFIG.ROUND;
  var stats={paresCompartida:0, paresExclusiva:0, paresPorCorreo:0};
  ventas.forEach(function(s){ s.round=false; s.roundMain=false; s.roundTipo=''; });

  var elegibles=[];
  ventas.forEach(function(s,i){ if(_concretadaR_(s.fin) && s.amount>0) elegibles.push(i); });

  // (1) COMPARTIDA: las dos piernas comparten el reservation_id; la vuelta trae sufijo Z
  var byBase={}, sinRid=[];
  elegibles.forEach(function(i){
    var s=ventas[i]; if(s.product!==CONFIG.PRODUCTO_COMPARTIDA) return;
    var rid=_ridNorm_(s.rid);
    if(!rid){ sinRid.push(i); return; }
    var b=_ridBase_(rid); (byBase[b]=byBase[b]||[]).push(i);
  });
  Object.keys(byBase).forEach(function(b){
    var g=byBase[b];
    if(g.length===2 && ventas[g[0]].agent===ventas[g[1]].agent){
      g.sort(function(x,y){ return ventas[x].date-ventas[y].date; });
      _marcarPar_(ventas[g[0]], ventas[g[1]], 'compartida'); stats.paresCompartida++;
    }
  });

  // (1b) respaldo: compartidas SIN reservation_id -> mismo agente + mismo correo, <=5 min
  if(sinRid.length){
    var idxC=[];
    elegibles.forEach(function(i){ if(ventas[i].product===CONFIG.PRODUCTO_COMPARTIDA) idxC.push(i); });
    idxC.sort(function(x,y){ var a=ventas[x], b=ventas[y];
      return (a.agent<b.agent?-1:a.agent>b.agent?1:0) || (a.user<b.user?-1:a.user>b.user?1:0) || (a.date-b.date); });
    var pos={}; idxC.forEach(function(i,k){ pos[i]=k; });
    sinRid.forEach(function(i){
      var s=ventas[i]; if(s.round||!s.user) return;
      for(var j=pos[i]-1;j>=0;j--){
        var p=ventas[idxC[j]];
        if(p.agent!==s.agent || p.user!==s.user) break;
        if((s.date-p.date)>CR.MIN_CREATED*MS_MIN) break;
        if(p.round) continue;
        _marcarPar_(p, s, 'compartida'); stats.paresCompartida++; stats.paresPorCorreo++; break;
      }
    });
  }

  // (2) EXCLUSIVA: mismo agente + mismo precio, <=5 min entre creaciones, tm_start con >24 h
  var byAgPr={};
  elegibles.forEach(function(i){
    var s=ventas[i]; if(s.product!==CONFIG.PRODUCTO_EXCLUSIVA) return;
    var k=s.agent+'|'+s.amount; (byAgPr[k]=byAgPr[k]||[]).push(i);
  });
  Object.keys(byAgPr).forEach(function(k){
    var g=byAgPr[k]; if(g.length<2) return;
    g.sort(function(x,y){ return ventas[x].date-ventas[y].date; });
    for(var a=0;a<g.length-1;a++){
      var A=ventas[g[a]]; if(A.round||!A.start) continue;
      for(var b=a+1;b<g.length;b++){
        var B=ventas[g[b]]; if(B.round||!B.start) continue;
        if((B.date-A.date) > CR.MIN_CREATED*MS_MIN) break;
        if(Math.abs(B.start-A.start) > CR.EXCL_MIN_START_H*MS_H){
          _marcarPar_(A,B,'exclusiva'); stats.paresExclusiva++; break;
        }
      }
    }
  });
  return stats;
}

/** Diagnostico: corre la deteccion sobre toda la base y devuelve el conteo. */
function diagnosticoRound(){
  var v=loadVentas_(), st=detectarRoundTrips_(v);
  var vr=0, mains=0; v.forEach(function(s){ if(s.round)vr++; if(s.roundMain)mains++; });
  var out={pares:mains, viajesRound:vr, detalle:st};
  Logger.log(JSON.stringify(out,null,2));
  return out;
}

function loadTurnosPrefijo_(prefijo){
  var _ck='TUR_'+norm_(prefijo);
  var _cached=_cacheGet_(_ck); if(_cached) return _cached;
  var ssT=ss_(CONFIG.TURNOS_SPREADSHEET_ID||CONFIG.VENTAS_SPREADSHEET_ID);
  var pref=norm_(prefijo), people={};
  ssT.getSheets().forEach(function(sh){
    if(norm_(sh.getName()).indexOf(pref)!==0) return;
    var v=sh.getDataRange().getValues(); if(!v.length) return;
    var dr=-1, best=0;
    for(var r=0;r<Math.min(8,v.length);r++){ var cnt=0;
      for(var c=1;c<v[r].length;c++) if(parseHeaderDate_(v[r][c])) cnt++;
      if(cnt>best){ best=cnt; dr=r; } }
    if(dr<0||best<3) return;
    var dateCols=[]; for(var c2=1;c2<v[dr].length;c2++) if(parseHeaderDate_(v[dr][c2])) dateCols.push(c2);
    for(var r2=dr+1;r2<v.length;r2++){
      var nm=v[r2][0]; if(!nm) continue; var s=String(nm).trim();
      if(!s||s==='.'||s==='Nombre'||s==='Cargo'||s==='Supervisor') continue;
      var key=nameKey_(s); var rec=people[key]||(people[key]={name:s, shifts:{}});
      dateCols.forEach(function(dc){
        var t=parseTurno_(v[r2][dc]); if(t.libre) return;
        var hd=parseHeaderDate_(v[dr][dc]); if(!hd) return;
        rec.shifts[dateKey_(hd)]=t;
      });
    }
  });
  _cachePut_(_ck, people, 300);
  return people;
}

function shiftHours_(start,end){ var out=[],h=start,off=0,g=0;
  while(h!==end && g<26){ out.push({off:off,h:h}); h++; if(h===24){h=0;off=1;} g++; } return out; }
function adminHoras_(start){ return CONFIG.ADMIN_POR_INGRESO[start] || [(start+4)%24,(start+5)%24,(start+6)%24]; }

// --------------------- DESEMPENO (Tableau crosstab) -------------------------
function tableauUrl_(path){ return CONFIG.TABLEAU_SERVER+'/api/'+CONFIG.TABLEAU_API_VERSION+path; }

function tableauSignIn_(){
  var resp=UrlFetchApp.fetch(tableauUrl_('/auth/signin'),{
    method:'post', contentType:'application/json',
    headers:{Accept:'application/json'},
    payload:JSON.stringify({credentials:{
      personalAccessTokenName:CONFIG.TABLEAU_PAT_NAME,
      personalAccessTokenSecret:CONFIG.TABLEAU_PAT_SECRET,
      site:{contentUrl:CONFIG.TABLEAU_SITE}}}),
    muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau signin fallo ('+resp.getResponseCode()+'): '+resp.getContentText().substring(0,300));
  var cred=JSON.parse(resp.getContentText()).credentials;
  return {token:cred.token, siteId:cred.site.id};
}

function tableauViewId_(auth){
  var filtro=encodeURIComponent('contentUrl:eq:'+CONFIG.TABLEAU_VIEW_CONTENT_URL);
  var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views?filter='+filtro),{
    headers:{'X-Tableau-Auth':auth.token, Accept:'application/json'}, muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau views fallo: '+resp.getContentText().substring(0,300));
  var vs=JSON.parse(resp.getContentText());
  var arr=(vs.views&&vs.views.view)||[];
  if(!arr.length) throw new Error('No se encontro la vista con contentUrl='+CONFIG.TABLEAU_VIEW_CONTENT_URL);
  return arr[0].id;
}

function parseNumLoc_(v){
  if(v==null||v==='') return null;
  var s=String(v).trim(); if(!s) return null;
  s=s.replace(/%/g,'').replace(/\s/g,'');
  if(/^-?\d{1,3}(\.\d{3})*,\d+$/.test(s)) s=s.replace(/\./g,'').replace(',','.');
  else if(/^-?\d+,\d+$/.test(s)) s=s.replace(',','.');
  var n=Number(s); return isNaN(n)?null:n;
}

/** Descarga la crosstab de Tableau y la vuelca en la pestana PERF_SHEET. */
function actualizarPerformance(){
  _bloquearSiRestringido_();
  var auth=tableauSignIn_();
  var viewId=tableauViewId_(auth);
  var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views/'+viewId+'/data?maxAge=5'),{
    headers:{'X-Tableau-Auth':auth.token}, muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau data fallo: '+resp.getContentText().substring(0,300));
  var txt=resp.getContentText().replace(/^\uFEFF/,'');
  var firstLine=txt.split(/\r?\n/)[0]||'';
  var delim=(firstLine.split(';').length>firstLine.split(',').length)?';':',';
  var rows=Utilities.parseCsv(txt, delim);
  if(rows.length<2) throw new Error('La crosstab llego vacia.');
  // mapear encabezados (tolerante a espacios y variantes)
  var H=rows[0].map(function(h){return String(h).trim().toLowerCase();});
  function col(part){ for(var i=0;i<H.length;i++) if(H[i].indexOf(part)>=0) return i; return -1; }
  var iE=col('email'), iN=col('fullname'),
      iFi=-1, iFu=col('full resolution'), iCs=col('csat'), iNp=col('nps');
  // FIRT: aceptar varios nombres posibles del encabezado
  ['first reply','reply time','first response','reply','respuesta','response','firt','frt','tmo respuesta'].forEach(function(p){ if(iFi<0) iFi=col(p); });
  // columna de # tickets / volumen resuelto (si existe en la crosstab)
  var iTk=-1;
  ['# tickets','tickets','volumen','resueltos','# casos','casos','count'].forEach(function(p){ if(iTk<0) iTk=col(p); });
  var iTkNum=-1;
  ['ticket number','ticket #','n° ticket','ticket'].forEach(function(p){ if(iTkNum<0) iTkNum=col(p); });
  // puede haber varias columnas con 'date' (ej. "Día de Date Time" y "Date Time"):
  // elegir la primera cuyos valores realmente parseen como fecha.
  var iD=-1;
  for(var ci=0;ci<H.length;ci++){
    if(H[ci].indexOf('date')<0 && H[ci].indexOf('fecha')<0) continue;
    for(var rr=1;rr<Math.min(rows.length,6);rr++){
      if(parseDate_(rows[rr][ci])){ iD=ci; break; }
    }
    if(iD>=0) break;
  }
  if(iE<0) throw new Error('La crosstab no trae columna Email. Encabezados: '+rows[0].join(' | '));
  // Unidades: si el encabezado dice (Min) se convierte a horas
  var fFi=(iFi>=0 && H[iFi].indexOf('(min')>=0)?(1/60):1;
  var fFu=(iFu>=0 && H[iFu].indexOf('(min')>=0)?(1/60):1;
  var out=[['Email','Nombre','Fecha','Firt_h','Furt_h','CSAT','NPS','Tickets','N_Ticket']];
  var mdyP=(iD>=0)?detectMDY_(rows, iD):false;
  for(var r=1;r<rows.length;r++){
    var em=String(rows[r][iE]||'').trim().toLowerCase(); if(!em||em.indexOf('@')<0) continue;
    var fecha='';                                  // sin columna de fecha = agregado de la vista
    if(iD>=0){ var d=parseDate_(rows[r][iD], mdyP); if(!d) continue; fecha=dateKey_(d); }
    var fi=iFi>=0?parseNumLoc_(rows[r][iFi]):null; if(fi!=null) fi=round_(fi*fFi,2);
    var fu=iFu>=0?parseNumLoc_(rows[r][iFu]):null; if(fu!=null) fu=round_(fu*fFu,2);
    var cs=iCs>=0?parseNumLoc_(rows[r][iCs]):null; if(cs!=null&&cs>1.5) cs=cs/100;   // 80 -> 0.80
    var np=iNp>=0?parseNumLoc_(rows[r][iNp]):null;
    var tk=iTk>=0?parseNumLoc_(rows[r][iTk]):null;
    // En esta crosstab cada agente trae 1 fila con las metricas y varias que solo listan
    // la fecha (metricas en blanco). Esas filas-fecha NO son tickets: se descartan.
    if(fi==null&&fu==null&&cs==null&&np==null&&tk==null) continue;
    out.push([em, iN>=0?rows[r][iN]:'', fecha, fi, fu, cs, np, tk, iTkNum>=0?String(rows[r][iTkNum]||'').trim():'']);
  }
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.PERF_SHEET)||book.insertSheet(CONFIG.PERF_SHEET);
  sh.clearContents();
  sh.getRange(1,1,out.length,9).setValues(out);
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  PropertiesService.getScriptProperties().setProperty(CONFIG.PERF_PROP, stamp);
  return {filas:out.length-1, actualizado:stamp, cols:{firt:iFi>=0, furt:iFu>=0, csat:iCs>=0, nps:iNp>=0, tickets:iTk>=0, fecha:iD>=0}, encabezados:rows[0], muestra:out.slice(0,4)};
}

/** Lee la pestana Performance (cache). */
function loadPerformance_(){
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.PERF_SHEET);
  if(!sh) return [];
  var v=sh.getDataRange().getValues(); if(v.length<2) return [];
  var rows=[];
  for(var i=1;i<v.length;i++){
    if(!v[i][0]) continue;
    var d=(v[i][2]!==''&&v[i][2]!=null)?parseDate_(v[i][2]):null;   // null = agregado sin fecha
    rows.push({email:String(v[i][0]).trim().toLowerCase(), nombre:v[i][1], date:d,
      firt:(v[i][3]===''||v[i][3]==null)?null:Number(v[i][3]),
      furt:(v[i][4]===''||v[i][4]==null)?null:Number(v[i][4]),
      csat:(v[i][5]===''||v[i][5]==null)?null:Number(v[i][5]),
      nps:(v[i][6]===''||v[i][6]==null)?null:Number(v[i][6]),
      tickets:(v[i][7]===''||v[i][7]==null)?null:Number(v[i][7]),
      ticketNum:(v[i][8]==null?'':String(v[i][8]))});
  }
  return rows;
}

// --------------------- AUDITORIAS (Tableau Support_CL_Audits) ---------------
function tableauViewByContentUrl_(auth, contentUrl){
  var filtro=encodeURIComponent('contentUrl:eq:'+contentUrl);
  var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views?filter='+filtro),{
    headers:{'X-Tableau-Auth':auth.token, Accept:'application/json'}, muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau views fallo: '+resp.getContentText().substring(0,300));
  var arr=((JSON.parse(resp.getContentText()).views)||{}).view||[];
  if(!arr.length) throw new Error('No se encontro la vista con contentUrl='+contentUrl);
  return arr[0].id;
}

/** Detecta si una columna de fechas viene mes-primero (M/D, formato US) o
 *  dia-primero (D/M). Devuelve true si es mes-primero. Mira valores que
 *  desambiguan: un componente >12 solo puede ser dia. ISO (yyyy-...) -> false. */
function detectMDY_(rows, ci){
  var dF=false, mF=false;
  for(var i=1;i<rows.length;i++){
    var dp=String(rows[i][ci]==null?'':rows[i][ci]).trim().split(/[ T]/)[0].split(/[\/\-]/);
    if(dp.length<2 || String(dp[0]).length===4) continue;
    var a=+dp[0], b=+dp[1];
    if(a>12) dF=true;     // primer componente es dia -> D/M
    if(b>12) mF=true;     // segundo componente es dia -> primero es mes -> M/D
  }
  return mF && !dF;       // mes-primero solo con evidencia y sin contradiccion
}

/** Descarga las auditorias de Tableau (Support_CL_Audits) a la pestana AUDIT_SHEET. */
function actualizarAuditorias(){
  _bloquearSiRestringido_();
  var auth=tableauSignIn_();
  var viewId=tableauViewByContentUrl_(auth, CONFIG.TABLEAU_AUDIT_CONTENT_URL);
  var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views/'+viewId+'/data?maxAge=5'),{
    headers:{'X-Tableau-Auth':auth.token}, muteHttpExceptions:true});
  if(resp.getResponseCode()>=300) throw new Error('Tableau data (auditorias) fallo: '+resp.getContentText().substring(0,300));
  var txt=resp.getContentText().replace(/^\uFEFF/,'');
  var firstLine=txt.split(/\r?\n/)[0]||'';
  var delim=(firstLine.split(';').length>firstLine.split(',').length)?';':',';
  var rows=Utilities.parseCsv(txt, delim);
  if(rows.length<2) throw new Error('La vista de auditorias llego vacia.');
  var H=rows[0].map(function(h){return String(h).trim().toLowerCase();});
  function col(){ for(var a=0;a<arguments.length;a++){ var part=arguments[a];
    for(var i=0;i<H.length;i++) if(H[i].indexOf(part)>=0) return i; } return -1; }
  var iEm=col('agent email','audited agent','email'),
      iTk=col('ticket'),
      iSc=col('total audit score','audit score','score');
  // columna de fecha: la primera (created/date/fecha/dttm) cuyos valores realmente parseen
  var iD=-1;
  for(var ci=0;ci<H.length;ci++){
    if(H[ci].indexOf('created')<0 && H[ci].indexOf('date')<0 && H[ci].indexOf('fecha')<0 && H[ci].indexOf('dttm')<0) continue;
    for(var rr=1;rr<Math.min(rows.length,6);rr++){ if(parseDate_(rows[rr][ci])){ iD=ci; break; } }
    if(iD>=0) break;
  }
  if(iEm<0||iSc<0) throw new Error('Faltan columnas de correo o score. Encabezados: '+rows[0].join(' | '));
  var mdyA=(iD>=0)?detectMDY_(rows, iD):false;
  var out=[['Email','Ticket','Fecha','Score']];
  for(var r=1;r<rows.length;r++){
    var em=String(rows[r][iEm]||'').trim().toLowerCase(); if(!em||em.indexOf('@')<0) continue;
    var sc=parseNumLoc_(rows[r][iSc]); if(sc==null) continue;
    var d=(iD>=0)?parseDate_(rows[r][iD], mdyA):null;
    out.push([em, iTk>=0?String(rows[r][iTk]||'').trim():'', d?dateKey_(d):'', sc]);
  }
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.AUDIT_SHEET)||book.insertSheet(CONFIG.AUDIT_SHEET);
  sh.clearContents();
  sh.getRange(1,1,out.length,4).setValues(out);
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  PropertiesService.getScriptProperties().setProperty(CONFIG.AUDIT_PROP, stamp);
  var ags={}; for(var q=1;q<out.length;q++) ags[out[q][0]]=1;
  return {filas:out.length-1, agentes:Object.keys(ags).length, actualizado:stamp};
}

/** Lee la pestana Auditorias (cache): [{email,ticket,date,score}]. */
function loadAuditorias_(){
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.AUDIT_SHEET);
  if(!sh) return [];
  var v=sh.getDataRange().getValues(); if(v.length<2) return [];
  var rows=[];
  for(var i=1;i<v.length;i++){
    if(!v[i][0]) continue;
    rows.push({email:String(v[i][0]).trim().toLowerCase(), ticket:String(v[i][1]||''),
      date:(v[i][2]!==''&&v[i][2]!=null)?parseDate_(v[i][2]):null, score:Number(v[i][3])});
  }
  return rows;
}

/** Re-sincroniza auditorias solo si la cache esta vacia o vieja (> AUDIT_MAX_AGE_H). */
function _autoRefreshAuditorias_(){
  var props=PropertiesService.getScriptProperties();
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.AUDIT_SHEET);
  var stamp=props.getProperty(CONFIG.AUDIT_PROP), stale=true;
  if(stamp){ var m=stamp.match(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/);
    if(m){ var t=new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5]); stale=((new Date()-t)/36e5)>CONFIG.AUDIT_MAX_AGE_H; } }
  // Solo baja de Tableau si la cache esta VACIA (primera vez). En cargas normales usa la cache
  // (rapido). Para refrescar: boton "Actualizar desde Tableau" o el trigger diario.
  if(!sh || sh.getLastRow()<2) actualizarAuditorias();
}

/** Re-sincroniza Performance (Tableau) solo si la cache esta vacia o vieja (> AUDIT_MAX_AGE_H). */
function _autoRefreshPerformance_(){
  var props=PropertiesService.getScriptProperties();
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.PERF_SHEET);
  var stamp=props.getProperty(CONFIG.PERF_PROP), stale=true;
  if(stamp){ var m=stamp.match(/(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})/);
    if(m){ var t=new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5]); stale=((new Date()-t)/36e5)>CONFIG.AUDIT_MAX_AGE_H; } }
  if(!sh || sh.getLastRow()<2) actualizarPerformance();
}

/** Diagnostico de las vistas de Tableau: imprime los encabezados reales y un
 *  conteo de filas/agentes. Ejecutar desde el editor y revisar el return o los Registros. */
function diagnosticoTableau(){
  var auth=tableauSignIn_(), res={};
  [['Performance',CONFIG.TABLEAU_VIEW_CONTENT_URL],['Auditorias',CONFIG.TABLEAU_AUDIT_CONTENT_URL]].forEach(function(par){
    try{
      var id=tableauViewByContentUrl_(auth, par[1]);
      var resp=UrlFetchApp.fetch(tableauUrl_('/sites/'+auth.siteId+'/views/'+id+'/data?maxAge=5'),{headers:{'X-Tableau-Auth':auth.token},muteHttpExceptions:true});
      var txt=resp.getContentText().replace(/^\uFEFF/,''), fl=txt.split(/\r?\n/)[0]||'';
      var delim=(fl.split(';').length>fl.split(',').length)?';':',';
      var rows=Utilities.parseCsv(txt, delim);
      res[par[0]]={filas:rows.length-1, encabezados:rows[0]};
    }catch(e){ res[par[0]]={error:e.message}; }
  });
  Logger.log(JSON.stringify(res,null,2));
  return res;
}

// --------------------- DATOS MAESTROS (planilla BUK) ------------------------
function rolDesdeEspecialidad_(esp){
  var e=norm_(esp);
  if(e.indexOf('EJECUTIV')>=0)    return 'Ejecutivo';
  if(e.indexOf('COORDINADOR')>=0) return 'Coordinador';
  if(e.indexOf('SUPERVISOR')>=0)  return 'Supervisor';
  if(e.indexOf('ANFITRI')>=0)     return 'Anfitrion';
  return 'Otro';
}
function tc_(s){ return String(s||'').toLowerCase().replace(/(^|\s)\S/g,function(c){return c.toUpperCase();}).trim(); }
function supCorto_(full){
  var t=String(full||'').trim().split(/\s+/).filter(Boolean);
  if(t.length>=4) return tc_(t[0]+' '+t[2]);   // 2 nombres + 2 apellidos
  if(t.length===3) return tc_(t[0]+' '+t[1]);
  return tc_(t.join(' '));
}

/** Recibe los registros parseados de la planilla BUK (hoja Trabajador) y
 *  sincroniza la pestana Maestros. regs: [{nombre, ap1, ap2, especialidad, supervisor}] */
function sincronizarMaestros(regs){
  _bloquearSiRestringido_();
  if(!regs||!regs.length) throw new Error('No llegaron registros de la planilla.');
  var book=rosterBook_();
  var sh=book.getSheetByName(CONFIG.MAESTROS_SHEET)||book.insertSheet(CONFIG.MAESTROS_SHEET);
  var prev={}, pv=sh.getDataRange().getValues();
  for(var i=1;i<pv.length;i++) if(pv[i][0]) prev[pv[i][0]]=1;
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  var out=[['NameKey','Nombre_Completo','Rol','Supervisor','Supervisor_Corto','Sincronizado']];
  var nuevos=0, vistos={};
  regs.forEach(function(r){
    var nombre=String(r.nombre||'').trim(), ap1=String(r.ap1||'').trim(), ap2=String(r.ap2||'').trim();
    if(!nombre||!ap1) return;
    var full=(nombre+' '+ap1+(ap2?' '+ap2:'')).replace(/\s+/g,' ');
    var key=nameKey_(nombre.split(/\s+/)[0]+' '+ap1);     // PRIMER nombre + PRIMER apellido = clave Turnos
    if(vistos[key]) return; vistos[key]=1;
    if(!prev[key]) nuevos++;
    out.push([key, tc_(full), rolDesdeEspecialidad_(r.especialidad), tc_(r.supervisor), supCorto_(r.supervisor), stamp]);
  });
  if(out.length<2) throw new Error('La planilla no trae filas validas (Nombre / Primer Apellido).');
  sh.clearContents();
  sh.getRange(1,1,out.length,6).setValues(out);
  PropertiesService.getScriptProperties().setProperty(CONFIG.MAESTROS_PROP, stamp);
  resetCacheDatos();   // datos maestros cambiaron: invalidar caché de lecturas
  return {total:out.length-1, nuevos:nuevos, actualizado:stamp};
}

/** Lee la pestana Maestros con estructuras para casar correos por nombre completo. */
function loadMaestros_(){
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.MAESTROS_SHEET);
  if(!sh) return [];
  var v=sh.getDataRange().getValues(), rows=[];
  for(var i=1;i<v.length;i++){
    if(!v[i][0]) continue;
    var toks=norm_(v[i][1]).split(/\s+/).filter(Boolean);
    var tokSet={}, pairSet={};
    toks.forEach(function(t){tokSet[t]=1;});
    for(var a=0;a<toks.length;a++) for(var b=0;b<toks.length;b++) if(a!==b) pairSet[toks[a]+toks[b]]=1;
    rows.push({key:String(v[i][0]), nombre:v[i][1], rol:v[i][2], supervisor:v[i][3], supCorto:v[i][4],
               tokSet:tokSet, pairSet:pairSet});
  }
  return rows;
}

// --------------------- AUSENTISMO / INCIDENCIAS (libro BSC) -----------------
/** Lee las incidencias de asistencia (hoja Consolidado del libro BSC) para los
 *  EJECUTIVOS, en el rango [fromDate,toDate], y las agrupa por nk (nameKey del
 *  nombre "360" de la hoja Alias, que coincide con el nk de comisiones).
 *  Devuelve {byNk, disponible, mensaje}. Nunca lanza: si no hay acceso al libro,
 *  disponible=false con un mensaje amable. */
function loadAusentismo_(fromDate, toDate){
  var res={byNk:{}, rutByNk:{}, nombreByNk:{}, disponible:false, mensaje:''};
  try{
    if(!CONFIG.BSC_SPREADSHEET_ID){ res.mensaje='No esta configurado el libro de asistencia (BSC).'; return res; }
    var bss=SpreadsheetApp.openById(CONFIG.BSC_SPREADSHEET_ID);
    var shCol=bss.getSheetByName(CONFIG.BSC_COLAB_SHEET);
    var shAli=bss.getSheetByName(CONFIG.BSC_ALIAS_SHEET);
    var shCon=bss.getSheetByName(CONFIG.BSC_CONSOLIDADO_SHEET);
    if(!shCol || !shCon){ res.mensaje='No se encontraron las hojas del libro de asistencia.'; return res; }

    // Colaboradores: id -> rol_operativo (para quedarnos solo con EJECUTIVO)
    var cv=shCol.getDataRange().getValues(), cc=getColMap_(cv[0]);
    var iId=cc['id'], iRol=cc['rol_operativo'], iRut=cc['rut'], iNom=cc['nombre_completo'];
    var rolById={}, rutById={}, nomById={};
    for(var i=1;i<cv.length;i++){ if(cv[i][iId]===''||cv[i][iId]==null) continue;
      rolById[String(cv[i][iId])]=norm_(cv[i][iRol]);
      rutById[String(cv[i][iId])]=(iRut!=null?String(cv[i][iRut]||'').trim():'');
      nomById[String(cv[i][iId])]=(iNom!=null?tc_(cv[i][iNom]):''); }

    // Alias (fuente 360): colaborador_id -> nk (nameKey del nombre 360 == nk de comisiones)
    var nkById={};
    if(shAli){
      var av=shAli.getDataRange().getValues(), ac=getColMap_(av[0]);
      var iCid=ac['colaborador_id'], iAlias=ac['alias'], iFuente=ac['fuente'];
      for(var a=1;a<av.length;a++){
        if(iFuente!=null && String(av[a][iFuente]||'').indexOf('360')<0) continue;   // solo alias de Turnos 360
        var cid=String(av[a][iCid]||''), alias=av[a][iAlias];
        if(cid && alias) nkById[cid]=nameKey_(alias);
      }
    }
    Object.keys(nkById).forEach(function(cid){ if(rolById[cid]==='EJECUTIVO'){ res.rutByNk[nkById[cid]]=rutById[cid]||''; res.nombreByNk[nkById[cid]]=nomById[cid]||''; } });

    var fromK=dateKey_(fromDate), toK=dateKey_(toDate);
    var DEDUCT={RETRASO:1, SALIDA_ANTICIPADA:1, INASISTENCIA_INJUSTIFICADA:1, MIXTO:1};
    var TRABAJABLE_BSC={CUMPLIDO:1, RETRASO:1, SALIDA_ANTICIPADA:1, INASISTENCIA_INJUSTIFICADA:1, MIXTO:1, SIN_MARCAJE_SALIDA:1};
    var OTRAS_BSC={SIN_MARCAJE_SALIDA:1, LICENCIA_MEDICA:1};
    var LABEL ={RETRASO:'Retraso', SALIDA_ANTICIPADA:'Salida anticipada',
      INASISTENCIA_INJUSTIFICADA:'Ausencia injustificada', INASISTENCIA_JUSTIFICADA:'Inasistencia justificada',
      MIXTO:'Retraso + salida anticipada', SIN_MARCAJE_SALIDA:'No marca salida', LICENCIA_MEDICA:'Licencia médica', PERMISO:'Permiso'};

    // (1) BSC / biometrico: "nk|dia" -> {tipo, min}
    var bscByKey={};
    var sv=shCon.getDataRange().getValues(), sc=getColMap_(sv[0]);
    var jId=sc['colaborador_id'], jF=sc['fecha'], jT=sc['tipo_calculado'],
        jMr=sc['minutos_retraso'], jMs=sc['minutos_salida_anticipada'],
        jHe=sc['hora_entrada_real'], jHs=sc['hora_salida_real'], jPi=sc['hora_inicio_plan'], jPf=sc['hora_fin_plan'];
    for(var r=1;r<sv.length;r++){
      var tc=String(sv[r][jT]||''); if(!tc) continue;
      var cid=String(sv[r][jId]||''); if(rolById[cid]!=='EJECUTIVO') continue;
      var f=sv[r][jF]; var fk=(f instanceof Date)? dateKey_(f) : (parseDate_(f)? dateKey_(parseDate_(f)) : '');
      if(!fk || fk<fromK || fk>toK) continue;
      var nk=nkById[cid]; if(!nk) continue;
      var mn=(tc==='RETRASO')?(Number(sv[r][jMr])||0):(tc==='SALIDA_ANTICIPADA')?(Number(sv[r][jMs])||0):(tc==='MIXTO')?((Number(sv[r][jMr])||0)+(Number(sv[r][jMs])||0)):0;
      bscByKey[nk+'|'+fk]={tipo:tc, min:round_(mn),
        entrada:(jHe!=null?_hhmm_(sv[r][jHe]):''), salida:(jHs!=null?_hhmm_(sv[r][jHs]):''),
        planIni:(jPi!=null?_hhmm_(sv[r][jPi]):''), planFin:(jPf!=null?_hhmm_(sv[r][jPf]):'')};
    }

    // (2) Manuales del supervisor (mandan sobre BSC): "nk|dia" -> {...}
    var man=loadIncidenciasManuales_(), manByKey={};
    Object.keys(man).forEach(function(k){ var day=k.substring(k.indexOf('|')+1); if(day>=fromK && day<=toK) manByKey[k]=man[k]; });

    // (3) Union de dias y calculo (el MANUAL gana sobre el automatico)
    var bucket=function(nk){ return res.byNk[nk]||(res.byNk[nk]={retrasos:0,ausencias:0,salidasAnt:0,mixto:0,sinSalida:0,licencia:0,
      retrasosDesc:0,ausenciasDesc:0,salidasDesc:0,mixtoDesc:0, manual:0,
      planned:0, deductEff:0, adherencia:null, detalle:[], disponible:true, mensaje:''}); };
    function countRaw(b,t){ if(t==='RETRASO')b.retrasos++; else if(t==='SALIDA_ANTICIPADA')b.salidasAnt++;
      else if(t==='INASISTENCIA_INJUSTIFICADA')b.ausencias++; else if(t==='MIXTO')b.mixto++;
      else if(t==='SIN_MARCAJE_SALIDA')b.sinSalida++; else if(t==='LICENCIA_MEDICA'||t==='INASISTENCIA_JUSTIFICADA')b.licencia++; }
    function countDesc(b,t){ if(t==='RETRASO')b.retrasosDesc++; else if(t==='SALIDA_ANTICIPADA')b.salidasDesc++;
      else if(t==='INASISTENCIA_INJUSTIFICADA')b.ausenciasDesc++; else if(t==='MIXTO')b.mixtoDesc++; }

    var allKeys={}; Object.keys(bscByKey).forEach(function(k){allKeys[k]=1;}); Object.keys(manByKey).forEach(function(k){allKeys[k]=1;});
    Object.keys(allKeys).forEach(function(key){
      var idx=key.indexOf('|'); var nk=key.substring(0,idx), day=key.substring(idx+1);
      var bsc=bscByKey[key], mm=manByKey[key];
      var b=bucket(nk);
      // contexto del biométrico para ese día (marcaje real + turno planificado), útil aunque el manual prevalezca
      var ctxE=(bsc?bsc.entrada:'')||'', ctxS=(bsc?bsc.salida:'')||'', ctxP=(bsc&&bsc.planIni)?(bsc.planIni+' → '+(bsc.planFin||'')):'';
      if(mm){   // ---- registro MANUAL (prevalece) ----
        b.manual++;
        var mtipo=mm.tipo, descuenta=(mm.descuenta==='si'), just=(mm.justificada==='si');
        var canon=null, trabajable=true, esInc=true, informativa=false;
        if(mtipo==='SIN_INCIDENCIA'){ esInc=false; }
        else if(mtipo==='INASISTENCIA'){ if(just){ canon='INASISTENCIA_JUSTIFICADA'; trabajable=false; informativa=true; descuenta=false; } else { canon='INASISTENCIA_INJUSTIFICADA'; } }
        else if(mtipo==='LICENCIA'){ canon='LICENCIA_MEDICA'; trabajable=false; informativa=true; descuenta=false; }
        else if(mtipo==='PERMISO'){ canon='PERMISO'; trabajable=false; informativa=true; descuenta=false; }
        else if(mtipo==='RETRASO'){ canon='RETRASO'; }
        else if(mtipo==='SALIDA_ANTICIPADA'){ canon='SALIDA_ANTICIPADA'; }
        else if(mtipo==='MIXTO'){ canon='MIXTO'; }
        else { esInc=false; }
        if(trabajable) b.planned++;
        if(esInc){
          countRaw(b, canon);
          if(descuenta && DEDUCT[canon]){ b.deductEff++; countDesc(b, canon); }
          b.detalle.push({fecha:day, tipo:canon, label:(LABEL[canon]||canon), min:round_(mm.min||0),
            descuenta:descuenta, editable:true, origen:'manual', informativa:informativa,
            justificada:(just?'si':'no'), comentario:(mm.comentario||''), supervisor:(mm.supervisor||''), tipoSel:mtipo,
            entrada:ctxE, salida:ctxS, turnoPlan:ctxP});
        } else {
          b.detalle.push({fecha:day, tipo:'SIN_INCIDENCIA', label:'Sin incidencia (revisado)', min:0,
            descuenta:false, editable:true, origen:'manual', informativa:true,
            justificada:'no', comentario:(mm.comentario||''), supervisor:(mm.supervisor||''), tipoSel:'SIN_INCIDENCIA',
            entrada:ctxE, salida:ctxS, turnoPlan:ctxP});
        }
      } else if(bsc){   // ---- solo BSC ----
        var tc2=bsc.tipo;
        if(TRABAJABLE_BSC[tc2]) b.planned++;
        if(DEDUCT[tc2]||OTRAS_BSC[tc2]){
          var d2=DEDUCT[tc2]?true:false;
          countRaw(b, tc2);
          if(d2 && DEDUCT[tc2]){ b.deductEff++; countDesc(b, tc2); }
          var tsel=(tc2==='RETRASO')?'RETRASO':(tc2==='SALIDA_ANTICIPADA')?'SALIDA_ANTICIPADA':(tc2==='INASISTENCIA_INJUSTIFICADA')?'INASISTENCIA':(tc2==='MIXTO')?'MIXTO':(tc2==='LICENCIA_MEDICA')?'LICENCIA':'SIN_INCIDENCIA';
          b.detalle.push({fecha:day, tipo:tc2, label:(LABEL[tc2]||tc2), min:round_(bsc.min||0),
            descuenta:d2, editable:(DEDUCT[tc2]?true:false), origen:'bsc', informativa:(OTRAS_BSC[tc2]?true:false),
            justificada:'no', comentario:'', supervisor:'', tipoSel:tsel,
            entrada:ctxE, salida:ctxS, turnoPlan:ctxP});
        }
      }
    });
    Object.keys(res.byNk).forEach(function(nk){ var b=res.byNk[nk];
      b.detalle.sort(function(a,b){return a.fecha<b.fecha?-1:(a.fecha>b.fecha?1:0);});
      b.adherencia = b.planned>0 ? round_((b.planned-b.deductEff)/b.planned,3) : null;
    });
    res.disponible=true;
  }catch(err){ res.disponible=false; res.mensaje=_amable_(err); }
  return res;
}

// --------------------- SUPERVISORES (auditan/sobrescriben incidencias) ------
function getSupervisores(){
  var raw=PropertiesService.getScriptProperties().getProperty(CONFIG.SUPERVISORES_PROP);
  if(!raw) return [];
  try{ var a=JSON.parse(raw); return Array.isArray(a)? a : []; }catch(e){ return []; }
}
function saveSupervisores(emails){
  _bloquearSiRestringido_();
  var arr=(emails||[]).map(function(x){ return String(x||'').trim().toLowerCase(); })
    .filter(function(x){ return x && x.indexOf('@')>0; });
  var seen={}, out=[]; arr.forEach(function(e){ if(!seen[e]){ seen[e]=1; out.push(e); } });
  PropertiesService.getScriptProperties().setProperty(CONFIG.SUPERVISORES_PROP, JSON.stringify(out));
  return out;
}

// --------------------- VISIBILIDAD DE COMISION -----------------------------
/** {agentes, coordinadores}: si true, el agente/coordinador ve su comision en su vista;
 *  si false, ve "Por definir". Por defecto ambos en OFF. */
function getMostrarComision(){
  var raw=PropertiesService.getScriptProperties().getProperty(CONFIG.MOSTRAR_COMISION_PROP);
  if(!raw) return {agentes:false, coordinadores:false};
  try{ var o=JSON.parse(raw); return {agentes:!!o.agentes, coordinadores:!!o.coordinadores}; }
  catch(e){ return {agentes:false, coordinadores:false}; }
}
function saveMostrarComision(agentes, coordinadores){
  _bloquearSiRestringido_();
  var obj={agentes:!!agentes, coordinadores:!!coordinadores};
  PropertiesService.getScriptProperties().setProperty(CONFIG.MOSTRAR_COMISION_PROP, JSON.stringify(obj));
  return obj;
}

// --------------------- OVERRIDES DE INCIDENCIAS -----------------------------
/** Lee las sobrescrituras: mapa "nk|fecha|tipo" -> {estado, comentario, supervisor, fecha}. */
function loadOverrides_(){
  var m={};
  try{
    var book=rosterBook_(), sh=book.getSheetByName(CONFIG.OVERRIDES_SHEET);
    if(!sh) return m;
    var v=sh.getDataRange().getValues(); if(v.length<2) return m;
    var c=getColMap_(v[0]);
    for(var i=1;i<v.length;i++){
      var nk=v[i][c['nk']], f=v[i][c['fecha']], t=v[i][c['tipo']];
      if(!nk||!f||!t) continue;
      var fk=(f instanceof Date)? dateKey_(f) : (parseDate_(f)? dateKey_(parseDate_(f)) : String(f).trim());
      m[String(nk).trim()+'|'+fk+'|'+String(t).trim()]={estado:String(v[i][c['estado']]||''),
        comentario:String(v[i][c['comentario']]||''), supervisor:String(v[i][c['supervisor']]||''),
        actualizado:String(v[i][c['actualizado']]||'')};
    }
  }catch(e){}
  return m;
}
/** Un supervisor autorizado cambia el estado de una incidencia (descuenta/no_descuenta) + comentario. */
function guardarOverrideIncidencia(nk, fecha, tipo, estado, comentario){
  var acc=getAccessContext();
  if(!acc.esSupervisor) throw new Error('Solo un supervisor autorizado puede modificar el estado de una incidencia.');
  nk=String(nk||'').trim(); fecha=String(fecha||'').trim(); tipo=String(tipo||'').trim();
  if(!nk||!fecha||!tipo) throw new Error('Faltan datos de la incidencia.');
  estado=(estado==='no_descuenta')?'no_descuenta':'descuenta';
  comentario=String(comentario||'').trim().substring(0,300);
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.OVERRIDES_SHEET);
  if(!sh){ sh=book.insertSheet(CONFIG.OVERRIDES_SHEET);
    sh.getRange(1,1,1,6).setValues([['nk','fecha','tipo','estado','comentario','supervisor']]);
    // columna extra 'actualizado' al final
    sh.getRange(1,7).setValue('actualizado'); }
  var v=sh.getDataRange().getValues(), c=getColMap_(v[0]);
  // la columna fecha se guarda como TEXTO para que Sheets no la convierta a tipo fecha
  try{ sh.getRange(1, c['fecha']+1, sh.getMaxRows(), 1).setNumberFormat('@'); }catch(e){}
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  var fila=-1;
  for(var i=1;i<v.length;i++){
    var sf=v[i][c['fecha']]; var sfk=(sf instanceof Date)?dateKey_(sf):(parseDate_(sf)?dateKey_(parseDate_(sf)):String(sf).trim());
    if(String(v[i][c['nk']]).trim()===nk && sfk===fecha && String(v[i][c['tipo']]).trim()===tipo){ fila=i+1; break; }
  }
  var row=[nk, fecha, tipo, estado, comentario, acc.email, stamp];
  if(fila>0) sh.getRange(fila,1,1,7).setValues([row]);
  else sh.appendRow(row);
  return {ok:true, estado:estado, comentario:comentario, supervisor:acc.email, actualizado:stamp};
}

// --------------------- INCIDENCIAS MANUALES (supervisor) --------------------
/** Lee las incidencias manuales: mapa "nk|fecha" -> {tipo, justificada, descuenta, min, comentario, supervisor, agente, actualizado}. */
function loadIncidenciasManuales_(){
  var m={};
  try{
    var book=rosterBook_(), sh=book.getSheetByName(CONFIG.MANUALES_SHEET);
    if(!sh) return m;
    var v=sh.getDataRange().getValues(); if(v.length<2) return m;
    var c=getColMap_(v[0]);
    for(var i=1;i<v.length;i++){
      var nk=v[i][c['nk']], f=v[i][c['fecha']], t=v[i][c['tipo']];
      if(!nk||!f||!t) continue;
      var fk=(f instanceof Date)? dateKey_(f) : (parseDate_(f)? dateKey_(parseDate_(f)) : String(f).trim());
      m[String(nk).trim()+'|'+fk]={tipo:String(t).trim().toUpperCase(),
        justificada:String(v[i][c['justificada']]||'').trim().toLowerCase(),
        descuenta:String(v[i][c['descuenta']]||'').trim().toLowerCase(),
        min:Number(v[i][c['minutos']])||0, comentario:String(v[i][c['comentario']]||''),
        supervisor:String(v[i][c['supervisor']]||''), agente:String(v[i][c['agente']]||''),
        actualizado:String(v[i][c['actualizado']]||'')};
    }
  }catch(e){}
  return m;
}
/** Registra/actualiza una incidencia manual (un registro por dia por agente). Solo supervisores. */
function guardarIncidenciaManual(nk, agente, fecha, tipo, justificada, descuenta, minutos, comentario){
  var acc=getAccessContext();
  if(!acc.esSupervisor) throw new Error('Solo un supervisor autorizado puede registrar incidencias.');
  nk=String(nk||'').trim(); fecha=String(fecha||'').trim(); tipo=String(tipo||'').trim().toUpperCase();
  if(!nk||!fecha||!tipo) throw new Error('Faltan datos: agente, fecha y tipo son obligatorios.');
  var TIPOS={RETRASO:1,SALIDA_ANTICIPADA:1,INASISTENCIA:1,MIXTO:1,SIN_INCIDENCIA:1,PERMISO:1,LICENCIA:1};
  if(!TIPOS[tipo]) throw new Error('Tipo de incidencia no valido.');
  justificada=(String(justificada||'').toLowerCase()==='si')?'si':'no';
  descuenta=(String(descuenta||'').toLowerCase()==='si')?'si':'no';
  minutos=Number(minutos)||0; comentario=String(comentario||'').trim().substring(0,300);
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.MANUALES_SHEET);
  if(!sh){ sh=book.insertSheet(CONFIG.MANUALES_SHEET);
    sh.getRange(1,1,1,9).setValues([['nk','agente','fecha','tipo','justificada','descuenta','minutos','comentario','supervisor']]);
    sh.getRange(1,10).setValue('actualizado'); }
  var v=sh.getDataRange().getValues(), c=getColMap_(v[0]);
  try{ sh.getRange(1, c['fecha']+1, sh.getMaxRows(), 1).setNumberFormat('@'); }catch(e){}
  var stamp=Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Santiago','dd-MM-yyyy HH:mm');
  var fila=-1;
  for(var i=1;i<v.length;i++){ var sf=v[i][c['fecha']]; var sfk=(sf instanceof Date)?dateKey_(sf):(parseDate_(sf)?dateKey_(parseDate_(sf)):String(sf).trim());
    if(String(v[i][c['nk']]).trim()===nk && sfk===fecha){ fila=i+1; break; } }
  var row=[nk, String(agente||''), fecha, tipo, justificada, descuenta, minutos, comentario, acc.email, stamp];
  if(fila>0) sh.getRange(fila,1,1,10).setValues([row]); else sh.appendRow(row);
  return {ok:true};
}
/** Elimina una incidencia manual (agente+fecha). Solo supervisores. */
function eliminarIncidenciaManual(nk, fecha){
  var acc=getAccessContext();
  if(!acc.esSupervisor) throw new Error('Solo un supervisor autorizado puede eliminar incidencias.');
  nk=String(nk||'').trim(); fecha=String(fecha||'').trim();
  var book=rosterBook_(), sh=book.getSheetByName(CONFIG.MANUALES_SHEET);
  if(!sh) return {ok:true};
  var v=sh.getDataRange().getValues(), c=getColMap_(v[0]);
  for(var i=v.length-1;i>=1;i--){ var sf=v[i][c['fecha']]; var sfk=(sf instanceof Date)?dateKey_(sf):(parseDate_(sf)?dateKey_(parseDate_(sf)):String(sf).trim());
    if(String(v[i][c['nk']]).trim()===nk && sfk===fecha){ sh.deleteRow(i+1); } }
  return {ok:true};
}

// ================== METAS Y TRAMOS POR PERIODO (YYYY-MM) ====================
// Se guardan por mes con VIGENCIA HACIA ADELANTE: un periodo sin configuracion
// propia usa la del ultimo periodo anterior configurado. Asi, definir agosto no
// altera como se analiza julio.

function _ymValido_(ym){ return /^\d{4}-\d{2}$/.test(String(ym||'')); }
function _ymHoy_(){ return dateKey_(new Date()).substring(0,7); }
function _ymDeFecha_(d){ return dateKey_(d).substring(0,7); }

/** Busca el valor vigente para ym: propio -> ultimo anterior -> default -> null. */
function _vigente_(map, ym){
  if(!map) return null;
  if(map[ym]) return {val:map[ym], origen:'propia', desde:ym};
  var keys=Object.keys(map).filter(function(k){ return k!=='default' && _ymValido_(k) && k<ym; }).sort();
  if(keys.length){ var k=keys[keys.length-1]; return {val:map[k], origen:'heredada', desde:k}; }
  if(map['default']) return {val:map['default'], origen:'inicial', desde:''};
  return null;
}

// ---------------------------- METAS ----------------------------------------
function _mapaMetas_(){
  var props=PropertiesService.getScriptProperties();
  var raw=props.getProperty(CONFIG.META_PERIODOS_PROP);
  if(raw){ try{ var o=JSON.parse(raw); if(o && typeof o==='object') return o; }catch(e){} }
  // migracion desde la meta unica antigua
  var legacy=props.getProperty(CONFIG.META_PROP), map={};
  if(legacy){ try{ var l=JSON.parse(legacy); map['default']={grupo:Number(l.grupo)||0}; }catch(e){} }
  return map;
}
function _guardarMapaMetas_(map){
  PropertiesService.getScriptProperties().setProperty(CONFIG.META_PERIODOS_PROP, JSON.stringify(map));
}
/** Meta grupal vigente para un periodo. Devuelve {grupo, origen, desde}. */
function getMetasPeriodo(ym){
  ym=_ymValido_(ym)?ym:_ymHoy_();
  var v=_vigente_(_mapaMetas_(), ym);
  if(!v) return {grupo:0, origen:'sin_definir', desde:''};
  return {grupo:Number(v.val.grupo)||0, origen:v.origen, desde:v.desde};
}
/** Compatibilidad: meta del mes en curso. */
function getMetas(){
  var m=getMetasPeriodo(_ymHoy_());
  var t=getTramosPeriodo(_ymHoy_());
  var t5e=t.ejec[t.ejec.length-1]||{inf:0}, t5c=t.coord[t.coord.length-1]||{inf:0};
  return {grupo:m.grupo, metaAgente:t5e.inf||0, metaCoordinador:t5c.inf||0};
}
/** Guarda la meta grupal SOLO para ese periodo (no toca los demas). */
function saveMetasPeriodo(ym, grupo){
  _bloquearSiRestringido_();
  if(!_ymValido_(ym)) throw new Error('Periodo invalido (formato esperado YYYY-MM).');
  var map=_mapaMetas_();
  map[ym]={grupo:Number(grupo)||0};
  _guardarMapaMetas_(map);
  return getMetasPeriodo(ym);
}
/** Borra la configuracion propia de metas de ese periodo (vuelve a heredar). */
function borrarMetasPeriodo(ym){
  _bloquearSiRestringido_();
  var map=_mapaMetas_(); delete map[ym]; _guardarMapaMetas_(map);
  return getMetasPeriodo(ym);
}

// ---------------------------- TRAMOS ---------------------------------------
function _tramosDefault_(){
  return {
    coord:[{inf:0,sup:18499999,pct:0.0},{inf:18500000,sup:25899999,pct:0.5},{inf:25900000,sup:33299999,pct:1.0},{inf:33300000,sup:40699999,pct:1.5},{inf:40700000,sup:null,pct:2.0}],
    ejec:[{inf:0,sup:2999999,pct:0.0},{inf:3000000,sup:4499999,pct:0.5},{inf:4500000,sup:5999999,pct:1.0},{inf:6000000,sup:7499999,pct:1.5},{inf:7500000,sup:null,pct:2.0}]
  };
}
function _normTramos_(arr){
  if(!arr||!arr.length) return null;
  return arr.map(function(t){ return {inf:Number(t.inf)||0,
    sup:(t.sup==null||t.sup===''?null:Number(t.sup)), pct:Number(t.pct)||0}; });
}
function _mapaTramos_(){
  var props=PropertiesService.getScriptProperties();
  var raw=props.getProperty(CONFIG.TRAMOS_PERIODOS_PROP);
  if(raw){ try{ var o=JSON.parse(raw); if(o && typeof o==='object') return o; }catch(e){} }
  // migracion desde los tramos unicos antiguos
  var legacy=props.getProperty(CONFIG.TRAMOS_PROP), map={};
  if(legacy){ try{ var l=JSON.parse(legacy);
    map['default']={coord:_normTramos_(l.coord)||_tramosDefault_().coord, ejec:_normTramos_(l.ejec)||_tramosDefault_().ejec}; }catch(e){} }
  return map;
}
function _guardarMapaTramos_(map){
  PropertiesService.getScriptProperties().setProperty(CONFIG.TRAMOS_PERIODOS_PROP, JSON.stringify(map));
}
/** Tramos vigentes para un periodo: {coord, ejec, origen, desde}. */
function getTramosPeriodo(ym){
  ym=_ymValido_(ym)?ym:_ymHoy_();
  var def=_tramosDefault_();
  var v=_vigente_(_mapaTramos_(), ym);
  if(!v) return {coord:def.coord, ejec:def.ejec, origen:'inicial', desde:''};
  return {coord:_normTramos_(v.val.coord)||def.coord, ejec:_normTramos_(v.val.ejec)||def.ejec,
          origen:v.origen, desde:v.desde};
}
/** Compatibilidad: tramos del mes en curso. */
function getTramos(){ var t=getTramosPeriodo(_ymHoy_()); return {coord:t.coord, ejec:t.ejec}; }
/** Guarda los tramos SOLO para ese periodo (no toca los demas). */
function saveTramosPeriodo(ym, coord, ejec){
  _bloquearSiRestringido_();
  if(!_ymValido_(ym)) throw new Error('Periodo invalido (formato esperado YYYY-MM).');
  var def=_tramosDefault_();
  var map=_mapaTramos_();
  map[ym]={coord:_normTramos_(coord)||def.coord, ejec:_normTramos_(ejec)||def.ejec};
  _guardarMapaTramos_(map);
  return getTramosPeriodo(ym);
}
function borrarTramosPeriodo(ym){
  _bloquearSiRestringido_();
  var map=_mapaTramos_(); delete map[ym]; _guardarMapaTramos_(map);
  return getTramosPeriodo(ym);
}

/** Todo lo que necesita la pestana de configuracion para un periodo. */
function getConfigPeriodo(ym){
  _bloquearSiRestringido_();
  ym=_ymValido_(ym)?ym:_ymHoy_();
  var m=getMetasPeriodo(ym), t=getTramosPeriodo(ym);
  var mm=_mapaMetas_(), mt=_mapaTramos_();
  var periodos={};
  Object.keys(mm).forEach(function(k){ if(_ymValido_(k)) (periodos[k]=periodos[k]||{}).metas=true; });
  Object.keys(mt).forEach(function(k){ if(_ymValido_(k)) (periodos[k]=periodos[k]||{}).tramos=true; });
  var lista=Object.keys(periodos).sort().map(function(k){
    return {ym:k, label:_labelPeriodo_(k), metas:!!periodos[k].metas, tramos:!!periodos[k].tramos}; });
  return {ym:ym, label:_labelPeriodo_(ym),
    metas:{grupo:m.grupo, origen:m.origen, desde:m.desde, desdeLabel:_labelPeriodo_(m.desde)},
    tramos:{coord:t.coord, ejec:t.ejec, origen:t.origen, desde:t.desde, desdeLabel:_labelPeriodo_(t.desde)},
    metaAgente:((t.ejec[t.ejec.length-1]||{}).inf)||0,
    metaCoordinador:((t.coord[t.coord.length-1]||{}).inf)||0,
    periodos:lista};
}
/** Borra metas + tramos propios de ese periodo (vuelve a heredar del anterior). */
function borrarConfigPeriodo(ym){
  _bloquearSiRestringido_();
  if(!_ymValido_(ym)) throw new Error('Periodo invalido.');
  borrarMetasPeriodo(ym); borrarTramosPeriodo(ym);
  return getConfigPeriodo(ym);
}
/** Copia la configuracion vigente de un periodo a otro (para "arrancar" un mes nuevo). */
function copiarConfigPeriodo(desde, hacia){
  _bloquearSiRestringido_();
  if(!_ymValido_(desde)||!_ymValido_(hacia)) throw new Error('Periodo invalido.');
  var m=getMetasPeriodo(desde), t=getTramosPeriodo(desde);
  saveMetasPeriodo(hacia, m.grupo);
  saveTramosPeriodo(hacia, t.coord, t.ejec);
  return getConfigPeriodo(hacia);
}
var MES_NOM_=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
function _labelPeriodo_(ym){
  if(!_ymValido_(ym)) return '';
  var p=ym.split('-'); return MES_NOM_[(+p[1])-1]+' '+p[0];
}

/** En que tramo cae un monto -> {tramo:1..n, pct}. */
function tramoDe_(monto, tabla){
  monto=Number(monto)||0;
  for(var i=0;i<tabla.length;i++){ var t=tabla[i];
    if(monto>=t.inf && (t.sup==null || monto<=t.sup)) return {tramo:i+1, pct:t.pct}; }
  var last=tabla[tabla.length-1]||{pct:0};
  return {tramo:tabla.length, pct:last.pct};
}
/** Comision con TOPE opcional: si el ultimo tramo tiene sup (no "en adelante"),
 *  la base comisionable se limita a ese tope para que la comision no escale infinito. */
function comisionDe_(monto, tabla){
  monto=Number(monto)||0;
  var td=tramoDe_(monto, tabla);
  var last=tabla[tabla.length-1]||{};
  var cap=(last.sup==null?null:last.sup);           // tope del ultimo tramo (null = "en adelante")
  var base=(cap!=null && monto>cap)? cap : monto;   // no escala mas alla del tope
  return {tramo:td.tramo, pct:td.pct, base:base, cap:cap, comision:round_(base*td.pct/100)};
}
/** %cumplimiento de un criterio de calidad: si real<inflexion => 0; si no => min(real/meta,1).
 *  real=null (sin dato) => 1 (100%, no se penaliza). Devuelve 0..1. */
function cumpl_(real, meta, inflexion){
  if(real==null || isNaN(real)) return 1;
  if(real < inflexion) return 0;
  var r = real/meta; return r>1?1:r;
}

// ------------------------------- DASHBOARD ---------------------------------
var MES_FULL_={enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,setiembre:8,octubre:9,noviembre:10,diciembre:11};
function sheetMonthYm_(name){
  var n=String(name).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'');
  for(var k in MES_FULL_){ if(n.indexOf(k)>=0) return CONFIG.ANIO+'-'+pad2_(MES_FULL_[k]+1); }
  return null;
}
/** roleKM["nameKey|YYYY-MM"] = rol, tomado de la hoja cuyo TITULO es ese mes.
 *  Prioridad ante conflicto del mismo mes: Supervisor > Coordinador > Ejecutivo > Anfitrion. */
function buildRoles_(){
  var _cached=_cacheGet_('ROLES'); if(_cached) return _cached;
  var ssT=ss_(CONFIG.TURNOS_SPREADSHEET_ID||CONFIG.VENTAS_SPREADSHEET_ID);
  var roleKM={}, nameByKey={};
  var RANK={Anfitrion:1, Ejecutivo:2, Coordinador:3, Supervisor:4};
  var PREF=[[CONFIG.PREFIJO_AGENTES,'Ejecutivo'],[CONFIG.PREFIJO_ANFITRIONES,'Anfitrion'],
            [CONFIG.PREFIJO_COORDINADORES,'Coordinador'],[CONFIG.PREFIJO_SUPERVISORES,'Supervisor']];
  ssT.getSheets().forEach(function(sh){
    var name=sh.getName(), rol=null;
    for(var i=0;i<PREF.length;i++){ if(norm_(name).indexOf(norm_(PREF[i][0]))===0){ rol=PREF[i][1]; break; } }
    if(!rol) return;
    var ym=sheetMonthYm_(name); if(!ym) return;
    var v=sh.getDataRange().getValues(); if(!v.length) return;
    var dr=-1,best=0; for(var r=0;r<Math.min(8,v.length);r++){ var cnt=0;
      for(var c=1;c<v[r].length;c++) if(parseHeaderDate_(v[r][c])) cnt++; if(cnt>best){best=cnt;dr=r;} }
    if(dr<0||best<3) return;
    for(var r2=dr+1;r2<v.length;r2++){
      var nm=v[r2][0]; if(!nm) continue; var s=String(nm).trim();
      if(!s||s==='.'||s==='Nombre'||s==='Cargo'||s==='Supervisor') continue;
      var has=false; for(var c3=1;c3<v[r2].length && !has;c3++){ if(!parseTurno_(v[r2][c3]).libre) has=true; }
      if(!has) continue;
      var key=nameKey_(s); nameByKey[key]=s;
      var cur=roleKM[key+'|'+ym];
      if(!cur || RANK[rol]>RANK[cur]) roleKM[key+'|'+ym]=rol;
    }
  });
  var _out={roleKM:roleKM, nameByKey:nameByKey};
  _cachePut_('ROLES', _out, 300);
  return _out;
}

function getDashboard(params){
  if(!_INTERNAL_BYPASS_) _bloquearSiRestringido_();   // agentes/coordinadores no acceden al dashboard completo
  params=params||{};
  var from=params.from?parseDate_(params.from):firstOfMonth_();
  var to=params.to?endOfDay_(parseDate_(params.to)):endOfDay_(new Date());
  var prodFilter=params.product||'todas', teamFilter=params.team||'todos';
  var withDetail=(params.incluirDetalle===true);   // los PDF piden el detalle por agente
  try{ _autoRefreshAuditorias_(); }catch(e){}       // auditorias al dia, sin boton manual
  try{ _autoRefreshPerformance_(); }catch(e){}      // desempeno al dia, sin boton manual

  // PERIODO DE CONFIGURACION: metas y tramos vigentes para el mes del corte (to).
  var ymCfg=_ymDeFecha_(to);
  var metasP=getMetasPeriodo(ymCfg);
  var tramosP=getTramosPeriodo(ymCfg);
  var tramos={coord:tramosP.coord, ejec:tramosP.ejec};

  var R=loadRoster_(), roster=R.byEmail, keyByEmail=R.keyByEmail;
  var ventas=loadVentas_();
  detectarRoundTrips_(ventas);        // marca round / roundMain en cada venta
  var agentPeople=loadTurnosPrefijo_(CONFIG.PREFIJO_AGENTES);
  var coordPeople=loadTurnosPrefijo_(CONFIG.PREFIJO_COORDINADORES);

  // ROL POR MES (temporal): el rol de cada mes se toma de la hoja cuyo TITULO es
  // ese mes (ej. el rol de mayo viene de "Agentes/Anfitriones/... Mayo"), no del
  // derrame de fechas de otra hoja. Asi un cambio de rol entre meses queda correcto.
  var RB=buildRoles_(), roleKM=RB.roleKM, nameByKey=RB.nameByKey;

  // Indices para casar correos SIN depender del Roster:
  //  - concatIdx: "CARLOSPALACIOS" -> "CARLOS PALACIOS" (nombre+apellido pegados en el correo)
  //  - firstIdx:  nombre de pila -> claves de Turnos (se usa solo si es UNICO, ej. "maribel")
  var concatIdx={}, firstIdx={};
  Object.keys(nameByKey).forEach(function(k){
    var p=k.split(' ');
    (firstIdx[p[0]]=firstIdx[p[0]]||[]).push(k);
    if(p.length>=2) concatIdx[p[0]+p[p.length-1]]=k;
  });

  // franjaCoords["date|hour"] = [coordName]  (horas activas, con wrap)
  // franjaShiftByNameHour["date|hour|coordName"] = {start,end}  (para atribuir la venta al tipo de turno)
  var franjaCoords={}, franjaShiftByNameHour={};
  Object.keys(coordPeople).forEach(function(key){ var p=coordPeople[key];
    Object.keys(p.shifts).forEach(function(dk){ var t=p.shifts[dk]; var base=parseDate_(dk); var admin=adminHoras_(t.start);
      shiftHours_(t.start,t.end).forEach(function(x){ if(admin.indexOf(x.h)>=0) return;
        var d=x.off?dateKey_(addDays_(base,1)):dk; (franjaCoords[d+'|'+x.h]=franjaCoords[d+'|'+x.h]||[]).push(p.name);
        franjaShiftByNameHour[d+'|'+x.h+'|'+p.name]={start:t.start, end:t.end}; }); }); });

  // Equipo del EJECUTIVO por su TURNO. Dos mapas:
  //  - agentDayTeam[nk|fechaInicioTurno] = equipo del turno que EMPIEZA ese dia
  //    (cubre TODAS sus ventas de ese dia, sin importar la hora exacta).
  //  - agentCover[nk|fecha|hora] = equipo que cubre esa hora (para la madrugada
  //    que pertenece a un turno nocturno iniciado el dia anterior).
  // cruza medianoche -> Noche (todo el turno); mismo dia -> Dia.
  var agentDayTeam={}, agentCover={};
  Object.keys(agentPeople).forEach(function(key){ var p=agentPeople[key];
    Object.keys(p.shifts).forEach(function(dk){ var t=p.shifts[dk]; var base=parseDate_(dk); var team=t.cross?'Noche':'Dia';
      agentDayTeam[key+'|'+dk]=team;
      shiftHours_(t.start,t.end).forEach(function(x){ var d=x.off?dateKey_(addDays_(base,1)):dk; agentCover[key+'|'+d+'|'+x.h]=team; }); }); });

  var COMP=CONFIG.PRODUCTO_COMPARTIDA, EXCL=CONFIG.PRODUCTO_EXCLUSIVA;
  var horas=[]; for(var h=0;h<24;h++) horas.push({hora:h,comp:0,excl:0,total:0,n:0});
  var dowSum=[0,0,0,0,0,0,0], dowN=[0,0,0,0,0,0,0];  // 0=Dom..6=Sab
  var kpis={totalVentas:0,nViajes:0,compMonto:0,exclMonto:0,compN:0,exclN:0,diaMonto:0,nocheMonto:0,sinMonto:0,
            roundPares:0,roundViajes:0,roundMonto:0};
  var ejec={}, coordCom={}, otras={}, malPorEjec={}, malCorreos=[];
  var ventasDet={}, malDet={}, coordVentasDet={}, coordTurnoSales={};   // detalle por nk / coordinador (solo cuando withDetail)
  var porDia={};                 // monto y viajes por fecha (serie diaria)
  var ultimaVenta=null;          // fecha del viaje mas reciente considerado (corte de datos)

  // DATOS MAESTROS (planilla BUK): permite casar CUALQUIER correo contra el
  // nombre completo (incluye 2dos nombres y 2dos apellidos) y da rol/supervisor de respaldo.
  var maestros=loadMaestros_(), mmCache={};
  function matchMaestros_(email){
    if(email in mmCache) return mmCache[email];
    var toks=norm_(String(email).split('@')[0]).split(/[._]/).filter(Boolean);
    var hit=null;
    if(toks.length){
      var hits=maestros.filter(function(m){
        return toks.every(function(t){ return m.tokSet[t]||m.pairSet[t]; });
      });
      if(hits.length===1) hit=hits[0];
    }
    mmCache[email]=hit; return hit;
  }

  var venCache={};
  function infoVendedor(email, ym){
    var ck=email+'|'+ym; if(venCache[ck]) return venCache[ck];
    var rinfo=roster[email];
    var mm=matchMaestros_(email);
    var t0=norm_(String(email).split('@')[0].replace(/_/g,'.').split('.').filter(Boolean)[0]||'');
    var cands=[];
    if(CONFIG.ALIAS_CORREOS[email]) cands.push(nameKey_(CONFIG.ALIAS_CORREOS[email])); // alias manual
    if(keyByEmail[email]) cands.push(keyByEmail[email]);                                // Roster (si existe)
    if(mm) cands.push(mm.key);                                                          // Maestros (planilla BUK)
    cands.push(keyFromEmail_(email));                                                   // nombre.apellido del correo
    if(concatIdx[t0]) cands.push(concatIdx[t0]);                                        // nombreapellido pegados
    if(firstIdx[t0] && firstIdx[t0].length===1) cands.push(firstIdx[t0][0]);            // nombre de pila unico
    var seen={}; cands=cands.filter(function(k){ if(!k||seen[k])return false; seen[k]=1; return true; });
    // nk para el equipo: el primer candidato que EXISTA en Turnos
    var nk=null; for(var i=0;i<cands.length;i++){ if(nameByKey[cands[i]]){ nk=cands[i]; break; } }
    if(!nk) nk=cands[0];
    // rol del periodo: Turnos del mes (cualquier candidato) -> Maestros -> Roster -> 'Otro'
    var rol=''; for(var j=0;j<cands.length;j++){ if(roleKM[cands[j]+'|'+ym]){ rol=roleKM[cands[j]+'|'+ym]; break; } }
    if(!rol) rol=(mm&&mm.rol)||(rinfo&&rinfo.rol)||'Otro';
    var nombre=(mm&&mm.nombre)||(rinfo&&rinfo.nombre)||nameByKey[nk]||email;
    var sup=(mm&&mm.supCorto)||(rinfo&&rinfo.supervisor)||'';
    var res={nk:nk, rol:rol, nombre:nombre, sup:sup};
    venCache[ck]=res; return res;
  }

  ventas.forEach(function(s){
    if(s.date<from||s.date>to) return;
    if(CONFIG.SOLO_CONCRETADAS && CONFIG.FINISH_CONCRETADAS.indexOf(s.fin)<0) return;
    var isComp=(s.product===COMP), isExcl=(s.product===EXCL);
    if(prodFilter==='compartida'&&!isComp) return;
    if(prodFilter==='exclusiva'&&!isExcl) return;

    var ven=infoVendedor(s.agent, s.dateKey.substring(0,7));
    var esEjec=(ven.rol==='Ejecutivo');

    // equipo Diurno/Nocturno SOLO aplica a ejecutivos (por su turno del dia).
    // Los no-ejecutivos no tienen equipo: no suman a Dia/Noche/Sin turno ni al grafico de franja.
    var team=null;
    if(esEjec){
      team=agentDayTeam[ven.nk+'|'+s.dateKey] || agentCover[ven.nk+'|'+s.dateKey+'|'+s.hour] || 'Sin turno';
    }
    if(teamFilter!=='todos' && team!==teamFilter) return;  // el filtro de equipo deja solo ejecutivos

    // Tarjetas / totales: TODAS las ventas (ejecutivos + otras)
    kpis.totalVentas+=s.amount; kpis.nViajes++;
    if(!ultimaVenta||s.date>ultimaVenta) ultimaVenta=s.date;
    var pd=porDia[s.dateKey]||(porDia[s.dateKey]={monto:0,n:0}); pd.monto+=s.amount; pd.n++;
    if(isComp){kpis.compMonto+=s.amount;kpis.compN++;}
    if(isExcl){kpis.exclMonto+=s.amount;kpis.exclN++;}
    if(s.round){ kpis.roundViajes++; kpis.roundMonto+=s.amount; if(s.roundMain) kpis.roundPares++; }
    // Franja horaria (TODAS las ventas) por compartida/exclusiva
    var hb=horas[s.hour]; hb.total+=s.amount; hb.n++;
    if(isComp)hb.comp+=s.amount; if(isExcl)hb.excl+=s.amount;
    // Dia de la semana (TODAS las ventas)
    var wd=s.date.getDay(); dowSum[wd]+=s.amount; dowN[wd]++;
    // Dia/Noche: SOLO ejecutivos (cards de equipo y tabla de ejecutivos)
    if(esEjec){
      if(team==='Noche')kpis.nocheMonto+=s.amount; else if(team==='Dia')kpis.diaMonto+=s.amount; else kpis.sinMonto+=s.amount;
    }

    // EJECUTIVOS (vendedor ejecutivo) u OTRAS VENTAS (vendedor no ejecutivo)
    if(esEjec){
      var e=ejec[ven.nk]||(ejec[ven.nk]={nombre:ven.nombre,sup:ven.sup,monto:0,n:0,comp:0,excl:0,compN:0,exclN:0,
                                         dia:0,noche:0,rPares:0,rViajes:0,rMonto:0});
      e.monto+=s.amount; e.n++; if(isComp){e.comp+=s.amount;e.compN++;} if(isExcl){e.excl+=s.amount;e.exclN++;}
      if(team==='Noche')e.noche+=s.amount; else if(team==='Dia')e.dia+=s.amount;
      if(s.round){ e.rViajes++; e.rMonto+=s.amount; if(s.roundMain) e.rPares++; }
      if(withDetail) (ventasDet[ven.nk]=ventasDet[ven.nk]||[]).push({fecha:s.dateKey, producto:s.product,
        monto:round_(s.amount), user:s.user, round:!!s.round, roundTipo:s.roundTipo||''});
    } else {
      var o=otras[s.agent]||(otras[s.agent]={nombre:ven.nombre,rol:ven.rol,monto:0,n:0});
      o.monto+=s.amount; o.n++;
    }

    // COMISION COORDINADORES por franja activa (50/50 si hay 2)
    var coords=franjaCoords[s.dateKey+'|'+s.hour]||[];
    if(coords.length){ var shr=s.amount/coords.length;
      coords.forEach(function(c){ var cc=coordCom[c]||(coordCom[c]={nombre:c,dia:0,noche:0,total:0,n:0});
        cc.total+=shr; cc.n+=1/coords.length; if(team==='Noche')cc.noche+=shr; else if(team==='Dia')cc.dia+=shr;
        if(withDetail){ (coordVentasDet[c]=coordVentasDet[c]||[]).push({fecha:s.dateKey, hora:s.hour,
          producto:s.product, montoVenta:round_(s.amount), montoCoord:round_(shr), nCoords:coords.length,
          vendedor:ven.nombre, rolVendedor:ven.rol, equipo:(team||'')});
          var _sh=franjaShiftByNameHour[s.dateKey+'|'+s.hour+'|'+c];
          if(_sh){ var _tk=_sh.start+'-'+_sh.end; var _ts=coordTurnoSales[c]||(coordTurnoSales[c]={}); _ts[_tk]=(_ts[_tk]||0)+shr; } } }); }

    // CORREOS MAL INGRESADOS (ds_user_email, solo compartida, vendedor ejecutivo)
    if(isComp&&esEjec){
      var mp=malPorEjec[ven.nk]||(malPorEjec[ven.nk]={nombre:ven.nombre,total:0,malos:0});
      mp.total++;
      if(!emailOk_(s.user)){ mp.malos++;
        var reg={ejecutivo:ven.nombre,correo:s.user,fecha:s.dateKey,monto:round_(s.amount)};
        if(malCorreos.length<3000) malCorreos.push(reg);
        if(withDetail) (malDet[ven.nk]=malDet[ven.nk]||[]).push(reg); }
    }
  });

  var t5e=tramos.ejec[tramos.ejec.length-1]||{inf:0}, t5c=tramos.coord[tramos.coord.length-1]||{inf:0};
  var metaAg=t5e.inf||0, metaCo=t5c.inf||0;   // meta individual = inicio del Tramo 5 del periodo

  // AUDITORIAS por agente (promedio). NO se filtra por el periodo de ventas: se usa una
  // ventana propia (CONFIG.AUDIT_VENTANA_DIAS; 0 = todas las auditorias en cache).
  var auditByNk={}, auditDetByNk={}, auditPeriodByNk={}, ymA=dateKey_(to).substring(0,7);
  var _audFromK=dateKey_(from), _audToK=dateKey_(to);
  var auDesde = (CONFIG.AUDIT_VENTANA_DIAS>0)? addDays_(to, -CONFIG.AUDIT_VENTANA_DIAS) : null;
  loadAuditorias_().forEach(function(a){
    if(auDesde && a.date && a.date<auDesde) return;
    if(a.score==null||isNaN(a.score)) return;
    var nk=infoVendedor(a.email, ymA).nk;
    var ag=auditByNk[nk]||(auditByNk[nk]={sum:0,n:0}); ag.sum+=a.score; ag.n++;
    // auditorias DEL PERIODO seleccionado (base del criterio de comision): sin estas, no se descuenta
    if(a.date){ var _adk=dateKey_(a.date); if(_adk>=_audFromK && _adk<=_audToK){
      var ap=auditPeriodByNk[nk]||(auditPeriodByNk[nk]={sum:0,n:0}); ap.sum+=a.score; ap.n++; } }
    if(withDetail) (auditDetByNk[nk]=auditDetByNk[nk]||[]).push({ticket:a.ticket, fecha:a.date?dateKey_(a.date):'', score:a.score});
  });

  // Ausentismo/incidencias por agente. Lo piden el dashboard (adherencia + comision real)
  // y la vista agente (detalle). La vista coordinador NO lo pide (no se lee el BSC ahi).
  var ausentismoByNk={}, ausDisp=false, ausMsg='', rutByNk={};
  if(params.incluirAusentismo===true){
    var _au=loadAusentismo_(from, to); ausentismoByNk=_au.byNk; ausDisp=_au.disponible; ausMsg=_au.mensaje; rutByNk=_au.rutByNk||{};
  }

  var ejecList=Object.keys(ejec).map(function(k){ var e=ejec[k];
    var au=auditByNk[k], mp=malPorEjec[k], cm=comisionDe_(e.monto, tramos.ejec);
    var _oneWay=(e.n-e.rViajes);
    var o={nombre:e.nombre, nk:k, sup:e.sup||'', ventas:round_(e.monto), n:e.n, comp:round_(e.comp), excl:round_(e.excl),
            compN:e.compN, exclN:e.exclN,
            dia:round_(e.dia), noche:round_(e.noche), meta:metaAg,
            avance: metaAg? round_(e.monto/metaAg,3):null,
            comisionPct:cm.pct, tramo:cm.tramo, comision:cm.comision,
            ventasRound:e.rPares, viajesRound:e.rViajes, montoRound:round_(e.rMonto), ventasOneWay:_oneWay,
            montoOneWay:round_(e.monto-e.rMonto),
            pctRound:((e.rPares+_oneWay)? round_(e.rPares/(e.rPares+_oneWay),3):null),
            notaAuditoria:(au&&au.n)? round_(au.sum/au.n,1):null, nAudit:(au?au.n:0),
            regTotal:(mp?mp.total:0), regMalos:(mp?mp.malos:0),
            pctOk:(mp&&mp.total)? round_(1-(mp.malos/mp.total),3):null};
    // --- Comision de calidad: base = comision de tramo; 70% fijo + 3 bonos de 10% ---
    var _aus = ausentismoByNk[k] || null;
    var _adher = _aus ? _aus.adherencia : null;
    var _cReg = cumpl_(o.pctOk, CONFIG.CALIDAD.registros.meta, CONFIG.CALIDAD.registros.inflexion);           // pctOk null => 100%
    var _ap = auditPeriodByNk[k];
    var _cAud = (_ap && _ap.n>0)? cumpl_((_ap.sum/_ap.n)/100, CONFIG.CALIDAD.auditoria.meta, CONFIG.CALIDAD.auditoria.inflexion) : 1;  // sin auditorias EN EL PERIODO => 100% (no se descuenta)
    o.nAuditPeriodo = _ap? _ap.n : 0;
    var _cAdh = (_adher!=null)? cumpl_(_adher, CONFIG.CALIDAD.adherencia.meta, CONFIG.CALIDAD.adherencia.inflexion) : 1;               // sin datos => 100%
    var _base = o.comision;
    o.rut = rutByNk[k] || '';
    o.adherencia=_adher; o.cReg=_cReg; o.cAud=_cAud; o.cAdh=_cAdh;
    o.comisionVenta  = round_(0.70*_base);
    o.bonoRegistros  = round_(_cReg*0.10*_base);
    o.bonoAuditoria  = round_(_cAud*0.10*_base);
    o.bonoAdherencia = round_(_cAdh*0.10*_base);
    o.comisionReal   = round_(o.comisionVenta + o.bonoRegistros + o.bonoAuditoria + o.bonoAdherencia);
    if(params.incluirAusentismo===true){
      o.ausentismo = _aus || {retrasos:0,ausencias:0,salidasAnt:0,mixto:0,sinSalida:0,licencia:0,planned:0,deductEff:0,adherencia:null,detalle:[],disponible:ausDisp,mensaje:ausMsg};
    }
    if(withDetail){ o.ventasDetalle=ventasDet[k]||[]; o.registrosMalos=malDet[k]||[]; o.auditDetalle=auditDetByNk[k]||[]; }
    return o; })
    .sort(function(a,b){return b.ventas-a.ventas;});

  // Info de turnos/franjas por coordinador. Se calcula SIEMPRE (es liviano): lo usan
  // tanto la tabla del dashboard como la Vista Coordinador.
  var coordShiftByName={};
  var fromK=dateKey_(from), toK=dateKey_(to);
  Object.keys(coordPeople).forEach(function(key){
    var p=coordPeople[key], name=p.name;
    var info=coordShiftByName[name]||(coordShiftByName[name]={turnos:0, franjasActivas:0, franjasSolo:0, franjasCompartidas:0, porTurno:{}});
    Object.keys(p.shifts).forEach(function(dk){
      if(dk<fromK||dk>toK) return;                 // turno cuyo inicio cae en el periodo
      var t=p.shifts[dk]; info.turnos++;           // turno NO libre (los libres no se guardan)
      var tk=t.start+'-'+t.end;                    // tipo de turno (por hora de inicio/fin)
      var bucket=info.porTurno[tk]||(info.porTurno[tk]={start:t.start, end:t.end, turnos:0, franjas:0});
      bucket.turnos++;
      var base=parseDate_(dk), admin=adminHoras_(t.start);
      shiftHours_(t.start,t.end).forEach(function(x){
        if(admin.indexOf(x.h)>=0) return;          // hora administrativa/colacion: no es franja activa
        var d=x.off?dateKey_(addDays_(base,1)):dk;
        info.franjasActivas++; bucket.franjas++;
        var nc=(franjaCoords[d+'|'+x.h]||[]).length;   // coordinadores activos esa hora
        if(nc>=2) info.franjasCompartidas++; else info.franjasSolo++;
      });
    });
  });

  var coordList=Object.keys(coordCom).map(function(c){ var x=coordCom[c]; var cm=comisionDe_(x.total, tramos.coord);
    var si=coordShiftByName[c]||{turnos:0,franjasActivas:0,franjasSolo:0,franjasCompartidas:0,porTurno:{}};
    var o={coordinador:c, dia:round_(x.dia), noche:round_(x.noche), total:round_(x.total), n:round_(x.n,1),
      meta:metaCo, avance: metaCo? round_(x.total/metaCo,3):null,
      comisionPct:cm.pct, tramo:cm.tramo, comision:cm.comision,
      turnos:si.turnos, franjasActivas:si.franjasActivas, franjasSolo:si.franjasSolo, franjasCompartidas:si.franjasCompartidas,
      ventaPromedioFranja: si.franjasActivas? round_(x.total/si.franjasActivas) : 0};
    if(withDetail){ o.ventasDetalle = coordVentasDet[c]||[];
      var _pt=(si.porTurno||{}), _tsc=coordTurnoSales[c]||{};
      o.turnosDetalle=Object.keys(_pt).map(function(tk){ var b=_pt[tk]; var venta=_tsc[tk]||0;
        return {turno:tk, start:b.start, end:b.end, turnos:b.turnos, franjas:b.franjas,
          venta:round_(venta), promedio: b.franjas? round_(venta/b.franjas):0}; })
        .sort(function(a,b){return b.venta-a.venta;}); }
    return o; })
    .sort(function(a,b){return b.total-a.total;});

  // RESUMEN del periodo (comisiones y cumplimiento)
  var comEjec=0, comCoord=0, agMeta=0, coMeta=0;
  ejecList.forEach(function(e){ comEjec+=e.comision; if(e.avance!=null&&e.avance>=1) agMeta++; });
  coordList.forEach(function(c){ comCoord+=c.comision; if(c.avance!=null&&c.avance>=1) coMeta++; });
  var resumen={comisionEjecutivos:round_(comEjec), comisionCoordinadores:round_(comCoord),
    comisionTotal:round_(comEjec+comCoord), nAgentes:ejecList.length, agentesEnMeta:agMeta,
    nCoordinadores:coordList.length, coordEnMeta:coMeta};

  var otrasList=Object.keys(otras).map(function(e){ return {email:e,nombre:otras[e].nombre,rol:otras[e].rol,
    monto:round_(otras[e].monto),n:otras[e].n}; }).sort(function(a,b){return b.monto-a.monto;});

  var malList=Object.keys(malPorEjec).map(function(k){ var m=malPorEjec[k]; return {nombre:m.nombre,
    total:m.total,malos:m.malos,pctOk:m.total?round_(1-(m.malos/m.total),3):null}; })
    .sort(function(a,b){return b.malos-a.malos;});

  // DESEMPENO (Tableau cache): se agrega TODA la vista publicada por agente; NO se
  // filtra por el periodo de ventas (la vista trae su propio rango y, como QA/Tableau
  // publican con desfase, filtrar por el mes de ventas dejaba el panel vacio).
  var ymRef=dateKey_(to).substring(0,7);
  var perf=loadPerformance_(), perfAgg={}, perfSinFechas=true, ticketsDetByNk={};
  perf.forEach(function(p){
    var a=perfAgg[p.email]||(perfAgg[p.email]={n:0,tk:0,fiOk:0,fin:0,fuOk:0,fun:0,cs:0,csn:0,np:0,npn:0});
    a.n++;
    if(p.tickets!=null) a.tk+=p.tickets;
    if(p.firt!=null){a.fin++; if(p.firt<=CONFIG.FIRT_UMBRAL_H)a.fiOk++;}
    if(p.furt!=null){a.fun++; if(p.furt<=CONFIG.FURT_UMBRAL_H)a.fuOk++;}
    if(p.csat!=null){a.cs+=p.csat;a.csn++;}
    if(p.nps!=null){a.np+=p.nps;a.npn++;}
    if(withDetail){ var pnk=infoVendedor(p.email, ymRef).nk;
      (ticketsDetByNk[pnk]=ticketsDetByNk[pnk]||[]).push({ticket:p.ticketNum||'', csat:p.csat, nps:p.nps, firt:p.firt, furt:p.furt}); }
  });
  if(withDetail) ejecList.forEach(function(e){ e.ticketsDetalle=ticketsDetByNk[e.nk]||[]; });
  var perfList=Object.keys(perfAgg).map(function(em){ var a=perfAgg[em]; var ven=infoVendedor(em, ymRef);
    return {nombre:ven.nombre, rol:ven.rol, tickets:a.n,   // cada registro de la vista = 1 ticket
      csat:a.csn?round_(a.cs/a.csn,3):null, nps:a.npn?round_(a.np/a.npn):null,
      firt:a.fin?round_(a.fiOk/a.fin,3):null,    // % cumplimiento primera respuesta <= 24h
      furt:a.fun?round_(a.fuOk/a.fun,3):null};   // % cumplimiento resolucion <= 120h
    })
    .sort(function(a,b){return (b.csat==null?-1:b.csat)-(a.csat==null?-1:a.csat);});
  var perfStamp=PropertiesService.getScriptProperties().getProperty(CONFIG.PERF_PROP)||'';

  // Promedio por dia de semana: suma del dia / cantidad de ese dia en el rango (hasta hoy)
  var diaCount=[0,0,0,0,0,0,0], hoy=new Date(), fin=(to<hoy?to:hoy);
  for(var dcur=new Date(from.getFullYear(),from.getMonth(),from.getDate()); dcur<=fin; dcur=addDays_(dcur,1)) diaCount[dcur.getDay()]++;
  var NOMBRE_DOW=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  var ORDEN=[1,2,3,4,5,6,0]; // Lunes..Domingo
  var semana=ORDEN.map(function(wd){ return {dia:NOMBRE_DOW[wd], total:round_(dowSum[wd]),
    dias:diaCount[wd], n:dowN[wd], promedio: diaCount[wd]? round_(dowSum[wd]/diaCount[wd]):0}; });

  // SERIE DIARIA: monto de ventas por fecha (todos los dias del rango, 0 si no hubo)
  var serieDiaria=[];
  for(var ds=new Date(from.getFullYear(),from.getMonth(),from.getDate()); ds<=to; ds=addDays_(ds,1)){
    var dk=dateKey_(ds), x=porDia[dk]||{monto:0,n:0};
    serieDiaria.push({fecha:dk, monto:round_(x.monto), n:x.n});
  }

  return {
    rango:{from:dateKey_(from),to:dateKey_(to)},
    corte:{ultima: ultimaVenta? (dateKey_(ultimaVenta)+' '+pad2_(ultimaVenta.getHours())+':'+pad2_(ultimaVenta.getMinutes())) : '', viajes:kpis.nViajes},
    kpis:{totalVentas:round_(kpis.totalVentas),nViajes:kpis.nViajes,compMonto:round_(kpis.compMonto),
      exclMonto:round_(kpis.exclMonto),compN:kpis.compN,exclN:kpis.exclN,
      diaMonto:round_(kpis.diaMonto),nocheMonto:round_(kpis.nocheMonto),sinMonto:round_(kpis.sinMonto),
      roundPares:kpis.roundPares, roundViajes:kpis.roundViajes, roundMonto:round_(kpis.roundMonto),
      oneWay:(kpis.nViajes-kpis.roundViajes)},
    franjas:horas.map(function(x){return {hora:x.hora,comp:round_(x.comp),excl:round_(x.excl),total:round_(x.total),n:x.n};}),
    semana:semana,
    serieDiaria:serieDiaria,
    desempeno:{lista:perfList, actualizado:perfStamp, sinFechas:perfSinFechas},
    auditoria:{actualizado: PropertiesService.getScriptProperties().getProperty(CONFIG.AUDIT_PROP)||''},
    maestros:{total:maestros.length, actualizado:PropertiesService.getScriptProperties().getProperty(CONFIG.MAESTROS_PROP)||''},
    ejecutivos:ejecList, coordinadores:coordList, coordShiftByName:coordShiftByName, otrasVentas:otrasList,
    correosMal:{detalle:malCorreos, porEjecutivo:malList},
    tramos:tramos,
    resumen:resumen,
    metas:{grupo:metasP.grupo||0, metaAgente:metaAg, metaCoordinador:metaCo},
    configPeriodo:{ym:ymCfg, label:_labelPeriodo_(ymCfg),
      metasOrigen:metasP.origen, metasDesde:metasP.desde, metasDesdeLabel:_labelPeriodo_(metasP.desde),
      tramosOrigen:tramosP.origen, tramosDesde:tramosP.desde, tramosDesdeLabel:_labelPeriodo_(tramosP.desde)},
    supervisores:getSupervisores(),
    mostrarComision:getMostrarComision()
  };
}

// ------------- helper de rango de mes para las vistas personales -----------
/** Devuelve {from, to, mesLabel, y, m} para 'actual' (mes en curso hasta hoy)
 *  o 'anterior' (mes anterior completo). */
function _rangoMes_(periodo){
  periodo=(periodo==='anterior')?'anterior':'actual';
  var n=new Date(), y=n.getFullYear(), m=n.getMonth();        // mes en curso (0-11)
  if(periodo==='anterior'){ m-=1; if(m<0){ m=11; y-=1; } }
  var desde=new Date(y, m, 1);
  var hasta=(periodo==='anterior') ? new Date(y, m+1, 0) : n;  // mes anterior completo / mes en curso hasta hoy
  var MES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return {from:dateKey_(desde), to:dateKey_(hasta), mesLabel:(MES[m]+' '+y), y:y, m:m, periodo:periodo};
}

// --------------- VISTA AGENTE (un agente, periodo seleccionable) -----------
/** Devuelve, para el periodo pedido ('actual' = mes en curso, 'anterior' = mes
 *  anterior completo), los ejecutivos con sus KPIs + detalle (ventas, registros
 *  incorrectos, tickets auditados). Si quien invoca es un AGENTE, la lista se
 *  filtra para que solo se vea a si mismo. */
function getMesAgentes(periodo){
  var R=_rangoMes_(periodo), from=R.from, to=R.to;
  var d;
  _INTERNAL_BYPASS_=true;
  try{ d=getDashboard({from:from, to:to, incluirDetalle:true, incluirAusentismo:true}); }
  catch(err){ _INTERNAL_BYPASS_=false; throw new Error(_amable_(err)); }
  finally { _INTERNAL_BYPASS_=false; }
  var perfByName={}; d.desempeno.lista.forEach(function(p){ perfByName[p.nombre]=p; });
  var ag=d.ejecutivos.map(function(e){ var p=perfByName[e.nombre]||{};
    // Auditorias SOLO del periodo seleccionado. Si no hay en el mes, no cuentan.
    var auds=(e.auditDetalle||[]).filter(function(a){ return a.fecha && a.fecha>=from && a.fecha<=to; });
    var nota=null; if(auds.length){ var sm=0; auds.forEach(function(a){ sm+=a.score; }); nota=Math.round(sm/auds.length*10)/10; }
    return {nombre:e.nombre, nk:e.nk, sup:e.sup, ventas:e.ventas, n:e.n, meta:e.meta, avance:e.avance,
      comisionPct:e.comisionPct, comision:e.comision, tramo:e.tramo,
      ventasRound:e.ventasRound, viajesRound:e.viajesRound, ventasOneWay:e.ventasOneWay,
      montoRound:e.montoRound, pctRound:e.pctRound,
      notaAuditoria:nota, nAudit:auds.length, regTotal:e.regTotal, regMalos:e.regMalos, pctOk:e.pctOk,
      tickets:(p.tickets!=null?p.tickets:null), csat:(p.csat!=null?p.csat:null), nps:(p.nps!=null?p.nps:null),
      firtPct:(p.firt!=null?p.firt:null), furtPct:(p.furt!=null?p.furt:null),
      ventasDetalle:e.ventasDetalle||[], registrosMalos:e.registrosMalos||[],
      auditDetalle:auds, ausentismo:e.ausentismo||{disponible:false,detalle:[]},
      adherencia:(e.adherencia==null?null:e.adherencia),
      cReg:e.cReg, cAud:e.cAud, cAdh:e.cAdh,
      comisionVenta:e.comisionVenta, bonoRegistros:e.bonoRegistros, bonoAuditoria:e.bonoAuditoria,
      bonoAdherencia:e.bonoAdherencia, comisionReal:e.comisionReal};
  });

  // CANDADO: un agente de venta solo se ve a si mismo
  var acc=getAccessContext();
  if(acc.esAgente){ ag=ag.filter(function(e){ return e.nombre===acc.nombre || e.nk===acc.nk; }); }

  return {rango:d.rango, periodo:R.periodo, mesLabel:R.mesLabel, corte:d.corte, configPeriodo:d.configPeriodo,
          ejecutivos:ag, mostrarComision:getMostrarComision(), acceso:{esAgente:acc.esAgente, esCoordinador:acc.esCoordinador, esSupervisor:acc.esSupervisor, nombre:acc.nombre, identificado:acc.identificado}};
}

// --------------- VISTA COORDINADOR (un coordinador, periodo selec.) --------
/** Igual que la vista de agente pero para COORDINADORES: sus ventas por franja
 *  (dia/noche/total), viajes, avance de meta y comision del periodo. Si quien
 *  invoca es un COORDINADOR, la lista se filtra para que solo se vea a si mismo. */
function getMesCoordinadores(periodo){
  var R=_rangoMes_(periodo), from=R.from, to=R.to;
  var d;
  _INTERNAL_BYPASS_=true;
  try{ d=getDashboard({from:from, to:to, incluirDetalle:true}); }
  catch(err){ _INTERNAL_BYPASS_=false; throw new Error(_amable_(err)); }
  finally { _INTERNAL_BYPASS_=false; }
  var lista=(d.coordinadores||[]).map(function(c){
    return {coordinador:c.coordinador, dia:c.dia, noche:c.noche, total:c.total, n:c.n,
      meta:c.meta, avance:c.avance, comisionPct:c.comisionPct, comision:c.comision, tramo:c.tramo,
      turnos:(c.turnos||0), franjasActivas:(c.franjasActivas||0), franjasSolo:(c.franjasSolo||0),
      franjasCompartidas:(c.franjasCompartidas||0), ventaPromedioFranja:(c.ventaPromedioFranja||0),
      ventasDetalle:c.ventasDetalle||[], turnosDetalle:c.turnosDetalle||[]};
  });
  // Coordinadores con turnos en el periodo pero SIN ventas: mostrarlos igual (turnos>0, ventas 0)
  var haveName={}; lista.forEach(function(c){ haveName[c.coordinador]=1; });
  var si=d.coordShiftByName||{};
  Object.keys(si).forEach(function(name){
    if(haveName[name]) return; var x=si[name]; if(!x.turnos) return;
    var tD=Object.keys(x.porTurno||{}).map(function(tk){ var b=x.porTurno[tk];
      return {turno:tk, start:b.start, end:b.end, turnos:b.turnos, franjas:b.franjas, venta:0, promedio:0}; });
    lista.push({coordinador:name, dia:0, noche:0, total:0, n:0,
      meta:(d.metas&&d.metas.metaCoordinador)||0, avance:((d.metas&&d.metas.metaCoordinador)?0:null),
      comisionPct:0, comision:0, tramo:1,
      turnos:x.turnos, franjasActivas:x.franjasActivas, franjasSolo:x.franjasSolo,
      franjasCompartidas:x.franjasCompartidas, ventaPromedioFranja:0, ventasDetalle:[], turnosDetalle:tD});
  });
  lista.sort(function(a,b){ return (b.total-a.total) || (b.turnos-a.turnos); });

  // CANDADO: un coordinador solo se ve a si mismo
  var acc=getAccessContext();
  if(acc.esCoordinador){
    lista=lista.filter(function(c){ return nameKey_(c.coordinador)===acc.nk || c.coordinador===acc.nombre; });
  }

  return {rango:d.rango, periodo:R.periodo, mesLabel:R.mesLabel, corte:d.corte, configPeriodo:d.configPeriodo,
          coordinadores:lista, mostrarComision:getMostrarComision(), acceso:{esAgente:acc.esAgente, esCoordinador:acc.esCoordinador, nombre:acc.nombre, identificado:acc.identificado}};
}

// --- Adherencia + comision real por agente para la tabla del dashboard ------
// Se pide en una llamada SEPARADA (en segundo plano) para no bloquear la carga
// principal del dashboard con la lectura del libro de asistencia (BSC).
function getAusentismoDashboard(params){
  _bloquearSiRestringido_();
  params=params||{}; params.incluirAusentismo=true;
  var d=getDashboard(params);
  return {disponible:true, agentes:(d.ejecutivos||[]).map(function(e){
    return {nombre:e.nombre, adherencia:(e.adherencia==null?null:e.adherencia), comisionReal:e.comisionReal};
  })};
}

// --------------------- VISTA PEOPLE (incidencias por periodo) ---------------
/** Devuelve el detalle y resumen de incidencias de los ejecutivos en un rango
 *  de fechas personalizable (para el equipo de People). Solo usuarios no restringidos. */
function getVistaPeople(desde, hasta){
  _bloquearSiRestringido_();
  var from=desde?parseDate_(desde):firstOfMonth_();
  var to=hasta?endOfDay_(parseDate_(hasta)):endOfDay_(new Date());
  if(!from) from=firstOfMonth_(); if(!to) to=endOfDay_(new Date());
  var au=loadAusentismo_(from, to);
  var filas=[], resumen={}, totalDesc=0;
  Object.keys(au.byNk||{}).forEach(function(nk){
    var b=au.byNk[nk];
    (b.detalle||[]).forEach(function(x){
      if(x.tipo==='SIN_INCIDENCIA') return;   // no es incidencia
      filas.push({colaborador:(au.nombreByNk[nk]||nk), rut:(au.rutByNk[nk]||''), fecha:x.fecha,
        incidencia:x.label, turnoPlan:(x.turnoPlan||''), entrada:(x.entrada||''), salida:(x.salida||''),
        min:(x.min||0), origen:(x.origen==='manual'?'Manual':'Biométrico'),
        efecto:(x.informativa?'Informativa':(x.descuenta?'Descuenta':'No descuenta')),
        comentario:(x.comentario||'')});
      resumen[x.label]=(resumen[x.label]||0)+1;
      if(!x.informativa && x.descuenta) totalDesc++;
    });
  });
  filas.sort(function(a,b){ return a.colaborador<b.colaborador?-1:a.colaborador>b.colaborador?1:(a.fecha<b.fecha?-1:a.fecha>b.fecha?1:0); });
  var resumenArr=Object.keys(resumen).map(function(k){ return {incidencia:k, total:resumen[k]}; }).sort(function(a,b){return b.total-a.total;});
  return {desde:dateKey_(from), hasta:dateKey_(to), filas:filas, resumen:resumenArr,
    totalIncidencias:filas.length, totalDescuentan:totalDesc, colaboradores:Object.keys(au.byNk||{}).length,
    disponible:au.disponible, mensaje:au.mensaje};
}

// ------------------------------- export CSV --------------------------------
function exportCSV(tipo,params){
  _bloquearSiRestringido_();
  params=params||{}; if(tipo==='ejecutivos') params.incluirAusentismo=true;   // el CSV de ejecutivos incluye adherencia y comision real
  var d=getDashboard(params), rows=[];
  if(tipo==='ejecutivos'){ rows.push(['Ejecutivo','RUT','Supervisor','Ventas','Viajes',
      'Q_Compartidas','Monto_Compartidas','Q_Exclusivas','Monto_Exclusivas','Ventas_Dia','Ventas_Noche',
      'Q_Round','Q_OneWay','Viajes_Round','Monto_Round','Monto_OneWay','%_Round',
      'Nota_Auditoria','N_Auditorias','%_Registros_OK','Registros_Total','Registros_Malos','%_Adherencia',
      'Meta','%_Avance','Tramo','%_Comision','Comision','Comision_Venta_70','Bono_Registros','Bono_Auditoria','Bono_Adherencia','Comision_Real']);
    d.ejecutivos.forEach(function(e){rows.push([e.nombre,(e.rut||''),e.sup,e.ventas,e.n,
      e.compN,e.comp,e.exclN,e.excl,e.dia,e.noche,
      e.ventasRound,e.ventasOneWay,e.viajesRound,e.montoRound,e.montoOneWay,e.pctRound,
      e.notaAuditoria,e.nAudit,e.pctOk,e.regTotal,e.regMalos,(e.adherencia==null?'':e.adherencia),
      e.meta,e.avance,e.tramo,e.comisionPct,e.comision,e.comisionVenta,e.bonoRegistros,e.bonoAuditoria,e.bonoAdherencia,e.comisionReal]);}); }
  else if(tipo==='coordinador'){ rows.push(['Coordinador','Ventas_Dia','Ventas_Noche','Ventas_Total','Viajes','Turnos','Prom_x_franja']);
    d.coordinadores.forEach(function(c){rows.push([c.coordinador,c.dia,c.noche,c.total,c.n,c.turnos,c.ventaPromedioFranja]);}); }
  else if(tipo==='grupo'){ rows.push(['Grupo','Ventas','% del total']);
    var tot=(d.kpis.diaMonto+d.kpis.nocheMonto+d.kpis.sinMonto)||1;
    rows.push(['Equipo Diurno',d.kpis.diaMonto,round_(d.kpis.diaMonto/tot,3)]);
    rows.push(['Equipo Nocturno',d.kpis.nocheMonto,round_(d.kpis.nocheMonto/tot,3)]);
    rows.push(['Sin turno',d.kpis.sinMonto,round_(d.kpis.sinMonto/tot,3)]);
    rows.push([]); rows.push(['Compartida',d.kpis.compMonto,d.kpis.compN]); rows.push(['Exclusiva',d.kpis.exclMonto,d.kpis.exclN]);
    rows.push([]); rows.push(['Ventas ida y vuelta',d.kpis.roundPares,d.kpis.roundViajes]); rows.push(['Monto ida y vuelta',d.kpis.roundMonto,'']);
    rows.push([]); rows.push(['Meta grupal',d.metas.grupo,d.metas.grupo?round_(d.kpis.totalVentas/d.metas.grupo,3):'']); }
  else if(tipo==='franja'){ rows.push(['Hora','Compartida','Exclusiva','Total','Viajes']);
    d.franjas.forEach(function(f){rows.push([f.hora,f.comp,f.excl,f.total,f.n]);}); }
  else if(tipo==='semana'){ rows.push(['Dia','Ventas_promedio','Total','Dias_en_periodo','Viajes']);
    d.semana.forEach(function(s){rows.push([s.dia,s.promedio,s.total,s.dias,s.n]);}); }
  else if(tipo==='diaria'){ rows.push(['Fecha','Ventas','Viajes']);
    d.serieDiaria.forEach(function(x){rows.push([x.fecha,x.monto,x.n]);}); }
  else if(tipo==='desempeno'){ rows.push(['Agente','Rol','Tickets','%_CSAT','NPS','%_FIRT_24h','%_FURT_120h']);
    d.desempeno.lista.forEach(function(p){rows.push([p.nombre,p.rol,p.tickets,p.csat,p.nps,p.firt,p.furt]);}); }
  else if(tipo==='otras'){ rows.push(['Email','Nombre','Rol','Ventas','Viajes']);
    d.otrasVentas.forEach(function(o){rows.push([o.email,o.nombre,o.rol,o.monto,o.n]);}); }
  else if(tipo==='correos'){ rows.push(['Ejecutivo','Correo_pasajero','Fecha','Monto']);
    d.correosMal.detalle.forEach(function(m){rows.push([m.ejecutivo,m.correo,m.fecha,m.monto]);}); }
  return rows.map(function(r){return r.map(csvCell_).join(',');}).join('\n');
}
function csvCell_(v){ v=(v==null?'':String(v)); return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; }

// ----------------------------- INFORMES PDF --------------------------------
// Construyen un Google Doc estilo Cabify (morado sobre blanco) y lo exportan a
// PDF en Drive. Devuelven {url, name}. Requiere scopes documents y drive
// (ver appsscript.json). El PDF queda en la carpeta CONFIG.AUDIT_PDF_FOLDER ('' = raiz).
var PURPLE_='#6C2EB9', PURPLE_DK_='#241046', LAV_='#B98BE8', MUT_='#8C84A3', INK_='#2B2440', LINE_='#ECE7F4';
function _fmtCLP_(n){ return '$'+Math.round(Number(n)||0).toLocaleString('es-CL'); }
function _pdfFolder_(){ return CONFIG.AUDIT_PDF_FOLDER? DriveApp.getFolderById(CONFIG.AUDIT_PDF_FOLDER) : DriveApp.getRootFolder(); }
function _hdr_(b, titulo, rango){
  b.appendParagraph('CABIFY · AEROPUERTO').editAsText().setForegroundColor(LAV_).setBold(true).setFontSize(9);
  var h=b.appendParagraph(titulo); h.setHeading(DocumentApp.ParagraphHeading.TITLE); h.editAsText().setForegroundColor(PURPLE_DK_);
  b.appendParagraph('Periodo: '+rango.from+'  a  '+rango.to).editAsText().setForegroundColor(MUT_).setFontSize(10);
  b.appendParagraph('');
}
function _sub_(b, txt){ var p=b.appendParagraph(txt); p.setHeading(DocumentApp.ParagraphHeading.HEADING2); p.editAsText().setForegroundColor(PURPLE_); }
function _nota_(b, txt){ b.appendParagraph(txt).editAsText().setForegroundColor(MUT_).setFontSize(10); }
function _tabla_(b, data){
  var t=b.appendTable(data); t.setBorderColor(LINE_);
  var hr=t.getRow(0);
  for(var c=0;c<hr.getNumCells();c++){ var cell=hr.getCell(c); cell.setBackgroundColor(PURPLE_);
    cell.editAsText().setForegroundColor('#FFFFFF').setBold(true).setFontSize(10); }
  for(var r=1;r<t.getNumRows();r++){ var row=t.getRow(r);
    for(var cc=0;cc<row.getNumCells();cc++) row.getCell(cc).editAsText().setBold(false).setFontSize(10).setForegroundColor(INK_); }
  return t;
}
function _docToPdf_(doc, nombre){
  doc.saveAndClose();
  var blob=DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  var b64=Utilities.base64Encode(blob.getBytes());
  DriveApp.getFileById(doc.getId()).setTrashed(true);   // no deja archivo suelto en Drive
  return {name:nombre+'.pdf', b64:b64};
}

/** Resumen de desempeño por agente: Ventas | Round | Nota Auditoría | % Registros Correctos. */
function generarResumenDesempenoPDF(params){
  _bloquearSiRestringido_();
  params=params||{}; var d=getDashboard(params);
  var titulo='Resumen de desempeño por agente';
  var doc=DocumentApp.create(titulo+' '+d.rango.from+' a '+d.rango.to);
  var b=doc.getBody(); b.setMarginTop(36).setMarginBottom(36).setMarginLeft(40).setMarginRight(40);
  _hdr_(b, titulo, d.rango);
  var data=[['Agente','Ventas','Ida y vuelta','Nota Auditoría','% Registros Correctos']];
  d.ejecutivos.forEach(function(e){
    data.push([e.nombre, _fmtCLP_(e.ventas),
      (e.ventasRound||0)+(e.pctRound!=null?(' ('+Math.round(e.pctRound*100)+'%)'):''),
      e.notaAuditoria==null?'—':String(e.notaAuditoria),
      e.pctOk==null?'—':(Math.round(e.pctOk*100)+'%')]); });
  if(data.length<2){ _nota_(b,'Sin ventas de ejecutivos en el periodo.'); }
  else _tabla_(b, data);
  return _docToPdf_(doc, titulo+' '+d.rango.from+'_'+d.rango.to);
}

/** Detalle de Desempeño por Agente: global + ventas + registros incorrectos + auditorias (ticket+nota).
 *  nombreAgente opcional: si se entrega, solo ese agente; si no, todos (uno por pagina). */
function generarDetalleDesempenoPDF(params, nombreAgente){
  _bloquearSiRestringido_();
  params=params||{}; params.incluirDetalle=true;
  var d=getDashboard(params);
  var lista=d.ejecutivos;
  if(nombreAgente) lista=lista.filter(function(e){return e.nombre===nombreAgente;});
  if(!lista.length) throw new Error('No hay datos de ejecutivos para el periodo'+(nombreAgente?(' ('+nombreAgente+')'):'')+'.');
  var titulo='Detalle de Desempeño por Agente';
  var doc=DocumentApp.create(titulo+' '+d.rango.from+' a '+d.rango.to);
  var b=doc.getBody(); b.setMarginTop(36).setMarginBottom(36).setMarginLeft(40).setMarginRight(40);
  _hdr_(b, titulo, d.rango);
  lista.forEach(function(e, idx){
    if(idx>0) b.appendPageBreak();
    var h=b.appendParagraph(e.nombre); h.setHeading(DocumentApp.ParagraphHeading.HEADING1); h.editAsText().setForegroundColor(PURPLE_);
    if(e.sup) _nota_(b,'Supervisor: '+e.sup);
    _sub_(b,'Resultado global');
    _tabla_(b, [['Ventas','Viajes','Ida y vuelta','Nota Auditoría','% Registros Correctos'],
      [_fmtCLP_(e.ventas), String(e.n),
       (e.ventasRound||0)+(e.pctRound!=null?(' ('+Math.round(e.pctRound*100)+'%)'):''),
       e.notaAuditoria==null?'—':String(e.notaAuditoria),
       e.pctOk==null?'—':(Math.round(e.pctOk*100)+'%')]]);
    _sub_(b,'Detalle de ventas del periodo ('+((e.ventasDetalle||[]).length)+')');
    var vt=[['Fecha','Servicio','Tipo','Monto','Correo pasajero']];
    (e.ventasDetalle||[]).forEach(function(v){ vt.push([v.fecha, v.producto, (v.round?'Ida y vuelta':'One way'), _fmtCLP_(v.monto), v.user||'—']); });
    if(vt.length>1) _tabla_(b, vt); else _nota_(b,'Sin ventas en el periodo.');
    _sub_(b,'Registros incorrectos ('+((e.registrosMalos||[]).length)+')');
    var rt=[['Fecha','Correo ingresado','Monto']];
    (e.registrosMalos||[]).forEach(function(m){ rt.push([m.fecha, m.correo||'—', _fmtCLP_(m.monto)]); });
    if(rt.length>1) _tabla_(b, rt); else _nota_(b,'Sin registros incorrectos.');
    _sub_(b,'Auditorías consideradas ('+((e.auditDetalle||[]).length)+')');
    var at=[['N° Ticket','Fecha ticket','Nota']];
    (e.auditDetalle||[]).forEach(function(a){ at.push([a.ticket||'—', a.fecha||'—', String(a.score)]); });
    if(at.length>1) _tabla_(b, at); else _nota_(b,'Sin auditorías registradas para el agente.');
  });
  return _docToPdf_(doc, titulo+(nombreAgente?(' '+nombreAgente):'')+' '+d.rango.from+'_'+d.rango.to);
}


// Ejecuta diagnostico() y revisa Registros (Ver > Registros). Te dice cuantas
// filas leyo del Roster y como queda clasificado un correo de ejemplo.
function diagnostico(){
  var book=rosterBook_();
  Logger.log('Libro del Roster: "'+book.getName()+'"  (hojas: '+book.getSheets().map(function(s){return s.getName();}).join(', ')+')');
  var R=loadRoster_();
  var emails=Object.keys(R.byEmail);
  Logger.log('Roster: '+emails.length+' correos leidos.');
  Logger.log('¿maribel.estefani@cabify.com en Roster? '+(R.byEmail['maribel.estefani@cabify.com']?('SI -> '+R.byEmail['maribel.estefani@cabify.com'].rol):'NO'));
  Logger.log('¿carlospalacios.alvarez@cabify.com en Roster? '+(R.byEmail['carlospalacios.alvarez@cabify.com']?('SI -> '+R.byEmail['carlospalacios.alvarez@cabify.com'].rol):'NO'));
  var ag=loadTurnosPrefijo_(CONFIG.PREFIJO_AGENTES);
  Logger.log('Agentes en Turnos 360: '+Object.keys(ag).length+' -> '+Object.keys(ag).slice(0,30).join(', '));
  return {rosterCount:emails.length, agentes:Object.keys(ag)};
}

/** Programa una actualizacion diaria de Tableau (auditorias + desempeno) ~06:00,
 *  para que la data quede fresca sin descargar de Tableau en cada carga del dashboard.
 *  Ejecutar UNA vez desde el editor (menu de funciones -> instalarActualizacionDiaria). */
function instalarActualizacionDiaria(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='_refrescoDiario_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('_refrescoDiario_').timeBased().everyDays(1).atHour(6).create();
  return 'Listo: Tableau se actualizara cada dia ~06:00. Las cargas del dashboard ya no bajaran de Tableau.';
}
function _refrescoDiario_(){ try{ actualizarAuditorias(); }catch(e){} try{ actualizarPerformance(); }catch(e){} }

/** Diagnostico de acceso: ejecutalo desde el editor logueado con tu cuenta para
 *  ver como te clasifica (o pruebalo mentalmente con distintos correos). */
function diagnosticoAcceso(){
  _ACC_CACHE_=null;
  var acc=getAccessContext();
  Logger.log(JSON.stringify(acc,null,2));
  return acc;
}
