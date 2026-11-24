const state = { data:null, selectedCampus:'__ALL__', selectedStudent:null, selectedPropensity:'all', selectedDashboardTab:'forecast', validationThreshold:0.35 };
const $ = id => document.getElementById(id);
const fmt = new Intl.NumberFormat('es-MX');
const pct = (v,d=1) => (v===null||v===undefined||Number.isNaN(Number(v))) ? '—' : `${(Number(v)*100).toFixed(d)}%`;
const pctAlready = (v,d=1) => (v===null||v===undefined||Number.isNaN(Number(v))) ? '—' : `${Number(v).toFixed(d)}%`;
const one = v => (v===null||v===undefined||Number.isNaN(Number(v))) ? '—' : Number(v).toFixed(1);
const safe = s => String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');
const cssState = s => s==='Riesgo de caída'?'risk':s==='Potencial de mejora'?'good':s==='Estable'?'good':s==='Seguimiento'?'warn':'neutral';
const scholarship = v => { if(v===null||v===undefined||Number.isNaN(Number(v)))return '—'; const n=Number(v); return `${(Math.abs(n)<=1.2?n*100:n).toFixed(0)}%`; };
const trafficClass = v => { if(v===null||v===undefined||Number.isNaN(Number(v))) return 'neutral'; const r=Math.round(Number(v)*10)/10; if(r<7)return 'red'; if(r<9)return 'yellow'; return 'green'; };
const trafficText = v => { if(v===null||v===undefined||Number.isNaN(Number(v))) return 'Sin dato'; const r=Math.round(Number(v)*10)/10; if(r<7)return 'Menor a 7'; if(r<9)return 'Entre 7 y 8.9'; return '9 o más'; };

const ANALYSIS_FILE = './base_continuidad_2027_dashboard.xlsx';
const BASE_VIVA_FILE = './BASE_VIVA.xlsx';
const INSCRITOS_FILE = './INSCRITOS_2027.xlsx';

const FACTOR_LABELS = {
  'factor_ambiente_estudiantil':'Ambiente estudiantil',
  'factor_atencion_personalizada_y_rapida':'Atención personalizada y rápida',
  'factor_costo_accesible':'Costo accesible',
  'factor_docentes_con_experiencia_academica_y_profesional':'Docentes con experiencia',
  'factor_instalaciones_de_primer_nivel':'Instalaciones de primer nivel',
  'factor_modelo_educativo_innovador':'Modelo educativo innovador',
  'factor_networking':'Networking',
  'factor_nivel_academico':'Nivel académico',
  'factor_oferta_de_becas':'Oferta de becas',
  'factor_oferta_de_carreras':'Oferta de carreras',
  'factor_oportunidades_de_intercambio_en_el_extranjero':'Intercambios en el extranjero',
  'factor_planes_de_estudios_innovadores':'Planes de estudio innovadores',
  'factor_podre_conseguir_un_buen_trabajo_al_graduarme':'Buen trabajo al graduarme',
  'factor_prestigio_de_la_universidad':'Prestigio de la universidad',
  'factor_prestigio_de_los_egresados':'Prestigio de los egresados',
  'factor_ubicacion':'Ubicación'
};
const REASON_LABELS = {
  'razon_me_da_mas_oportunidades_egresar_en_otra_universidad':'Más oportunidades al egresar en otra universidad',
  'razon_me_dan_mas_beca_en_otra_universidad':'Más beca en otra universidad',
  'razon_me_gusta_mas_el_plan_de_estudios_de_otra_universidad':'Me gusta más el plan de estudios de otra universidad',
  'razon_me_queda_lejos':'Me queda lejos',
  'razon_me_voy_a_ir_a_estudiar_en_el_extranjero':'Estudiar en el extranjero',
  'razon_mis_amigos_no_se_van_a_inscribir':'Mis amigos no se van a inscribir',
  'razon_no_hay_equipo_representativo_del_deporte_que_practico':'No hay equipo representativo de mi deporte',
  'razon_no_me_dan_beca_en_tecmilenio':'No me dan beca en Tecmilenio',
  'razon_no_tiene_la_carrera_que_busco':'No tiene la carrera que busco',
  'razon_por_influencia_de_algun_amigo':'Influencia de algún amigo',
  'razon_por_influencia_de_mis_papas':'Influencia de mis papás',
  'razon_recibi_mejor_atencion_durante_mi_proceso_de_admision_en_otra_universidad':'Mejor atención en otra universidad',
  'razon_vivencia_estudiantil':'Vivencia estudiantil'
};
const PROP_ORDER=['Muy baja','Baja','Media','Alta','Muy alta'];
const OPPORTUNITY_LEVELS=new Set(['Muy alta','Alta','Media-Alta']);

const cleanText = v => v===null||v===undefined ? '' : String(v).replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
const normalizeId = v => cleanText(v).toUpperCase().replace(/\.0+$/,'').replace(/[^A-Z0-9]/g,'');
const normalizeCampus = v => {
  let s=cleanText(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
  const aliases={'cd. juarez':'ciudad juarez','cd juarez':'ciudad juarez','cd. obregon':'obregon','cd obregon':'obregon','ciudad obregon':'obregon','san nicolas de los garza':'san nicolas'};
  return aliases[s]||s;
};
const displayCampus = v => {
  const n=normalizeCampus(v);
  const names={'cancun':'Cancún','chihuahua':'Chihuahua','ciudad juarez':'Ciudad Juárez','cuautitlan izcalli':'Cuautitlán Izcalli','cuernavaca':'Cuernavaca','culiacan':'Culiacán','cumbres':'Cumbres','durango':'Durango','ferreria':'Ferrería','guadalajara':'Guadalajara','guadalupe':'Guadalupe','hermosillo':'Hermosillo','laguna':'Laguna','las torres':'Las Torres','los mochis':'Los Mochis','mazatlan':'Mazatlán','merida':'Mérida','morelia':'Morelia','nuevo laredo':'Nuevo Laredo','obregon':'Obregón','puebla':'Puebla','queretaro':'Querétaro','reynosa':'Reynosa','san luis potosi':'San Luis Potosí','san nicolas':'San Nicolás','toluca':'Toluca','veracruz':'Veracruz','villahermosa':'Villahermosa','zacatecas':'Zacatecas','zapopan':'Zapopan'};
  return names[n] || (cleanText(v) ? cleanText(v).replace(/\b\w/g,c=>c.toUpperCase()) : 'Sin campus');
};
const asFloat = v => {
  if(v===null||v===undefined||v==='') return null;
  const n=Number(v); return Number.isFinite(n)?n:null;
};
const asInt = v => { const n=asFloat(v); return n===null?null:Math.round(n); };
const sheetNorm = s => cleanText(s).toLowerCase().replace(/[_\s]+/g,' ');

async function readWorkbook(url,reload=false){
  if(typeof XLSX==='undefined') throw new Error('No se pudo cargar el lector de Excel (SheetJS). Revisa tu conexión a internet.');
  const sep=url.includes('?')?'&':'?';
  const res=await fetch(`${url}${sep}v=${reload?Date.now():'1'}`,{cache:reload?'no-store':'default'});
  if(!res.ok) throw new Error(`No se pudo leer ${url.split('/').pop()} (${res.status}).`);
  const buf=await res.arrayBuffer();
  const wb=XLSX.read(buf,{type:'array',cellDates:false});
  const out={};
  wb.SheetNames.forEach(name=>{ out[name]=XLSX.utils.sheet_to_json(wb.Sheets[name],{defval:null,raw:true}); });
  return out;
}
function findSheet(book,...candidates){
  const map={}; Object.keys(book).forEach(k=>map[sheetNorm(k)]=k);
  for(const c of candidates){ const k=map[sheetNorm(c)]; if(k) return book[k]; }
  return [];
}
function weightedHist(rows,years){
  let nume=0,den=0;
  rows.forEach(r=>years.forEach(y=>{ const rate=asFloat(r[`tasa_${y}`]), n=asFloat(r[`N_${y}`]); if(rate!==null&&n!==null&&n>0){nume+=rate*n;den+=n;} }));
  return den?nume/den:null;
}
function historySummary(rows){
  const out=[];
  for(let y=2021;y<=2026;y++){
    let nume=0,den=0;
    rows.forEach(r=>{ const rate=asFloat(r[`tasa_${y}`]), n=asFloat(r[`N_${y}`]); if(rate!==null&&n!==null&&n>0){nume+=rate*n;den+=n;} });
    out.push({year:y,n:den?Math.round(den):0,rate:den?nume/den:null,continuers:den?Math.round(nume):0});
  }
  return out;
}
function average(rows,key){ const vals=rows.map(r=>asFloat(r[key])).filter(v=>v!==null); return vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null; }
function countBinary(rows,key){ return rows.reduce((n,r)=>n+(asFloat(r[key])===1?1:0),0); }
function groupBy(rows,keyFn){ const m=new Map(); rows.forEach(r=>{const k=keyFn(r); if(!m.has(k))m.set(k,[]); m.get(k).push(r);}); return m; }

function normalizeStatusValue(v){
  return cleanText(v).trim().toLowerCase();
}
function statusTokens(v){
  return normalizeStatusValue(v).split('/').map(x=>x.trim()).filter(Boolean);
}
function isEnrolledStatus(v){
  return statusTokens(v).includes('inscrito');
}
function isProspectStatus(v){
  const tokens=statusTokens(v);
  return !tokens.includes('inscrito') && tokens.includes('prospecto');
}

async function fetchData(reload=false){
  // En GitHub Pages los Excel se leen directamente en el navegador.
  // Si el lector XLSX no está disponible (por ejemplo, sin internet al abrir localmente),
  // el lanzador local puede usar la API del servidor Python como respaldo.
  if(typeof XLSX==='undefined'){
    const api=reload?`/api/reload?t=${Date.now()}`:`/api/data?t=${Date.now()}`;
    const res=await fetch(api,{cache:'no-store'});
    if(res.ok){ const json=await res.json(); if(json.ok) return json.data; }
    throw new Error('No se pudo cargar el lector de Excel. En GitHub Pages revisa la conexión; en local usa abrir_tablero.bat.');
  }
  const [analysis,vivaBook,inscritosBook]=await Promise.all([
    readWorkbook(ANALYSIS_FILE,reload),
    readWorkbook(BASE_VIVA_FILE,reload),
    readWorkbook(INSCRITOS_FILE,reload)
  ]);

  const alumnosAll=findSheet(analysis,'Alumnos');
  const historico=findSheet(analysis,'Historico campus','Historico_Campus');
  const catalogo=findSheet(analysis,'Catalogo_clusters','Catalogo Clusters');
  const vivaRows=vivaBook[Object.keys(vivaBook)[0]]||[];

  let inscritosRows=findSheet(inscritosBook,'Inscritos_2027','Inscritos 2027','Inscritos');
  if(!inscritosRows.length) inscritosRows=inscritosBook[Object.keys(inscritosBook)[0]]||[];

  if(!alumnosAll.length) throw new Error('El archivo analítico no contiene la hoja Alumnos o está vacía.');
  if(!vivaRows.length) throw new Error('La Base Viva está vacía.');
  if(!inscritosRows.length) throw new Error('INSCRITOS_2027.xlsx está vacío o no se pudo identificar su hoja.');

  // =====================================================
  // 1) BASE VIVA
  // Se mantiene EXACTAMENTE como en la versión anterior.
  // Representa a quienes siguen actualmente en el funnel.
  // =====================================================
  const vivaById=new Map();
  vivaRows.forEach(r=>{
    const mid=normalizeId(
      r['Matrícula'] ??
      r['Matricula'] ??
      r['Número de matrícula'] ??
      r['Numero de matricula']
    );
    if(!mid)return;

    const campus=r['Campus']??r['Escuela: Campus: Nombre de la cuenta']??'';
    const estatus=cleanText(r['Estatus']??r['ESTATUS']??r['estatus']);

    if(!vivaById.has(mid)){
      vivaById.set(mid,{
        campus:displayCampus(campus),
        campus_norm:normalizeCampus(campus),
        estatus:estatus
      });
    }else{
      const existing=vivaById.get(mid);
      if(!existing.campus && campus){
        existing.campus=displayCampus(campus);
        existing.campus_norm=normalizeCampus(campus);
      }
      if(estatus && !existing.estatus){
        existing.estatus=estatus;
      }else if(estatus && existing.estatus && !existing.estatus.split(' / ').includes(estatus)){
        existing.estatus=`${existing.estatus} / ${estatus}`;
      }
    }
  });

  // =====================================================
  // 2) RESPONDIENTES
  // Padrón maestro de este dashboard = hoja Alumnos.
  // =====================================================
  const respondentById=new Map();
  alumnosAll.forEach(r=>{
    const mid=normalizeId(r['Matrícula']??r['Matricula']);
    if(!mid)return;

    const campus=r['Campus']??'';
    if(!respondentById.has(mid)){
      respondentById.set(mid,{
        row:r,
        campus:displayCampus(campus),
        campus_norm:normalizeCampus(campus)
      });
    }
  });

  // =====================================================
  // 3) INSCRITOS CONFIRMADOS DE ESTA COHORTE
  // Regla EXACTA:
  // respondientes (Alumnos) ∩ INSCRITOS_2027 por matrícula.
  //
  // El archivo INSCRITOS_2027 puede traer Preparatoria,
  // Profesional u otros registros. Solo se cuentan aquí
  // quienes tienen matrícula en Alumnos.
  // =====================================================
  const inscritosSourceIds=new Set();
  const inscritosMatchedById=new Map();

  inscritosRows.forEach(r=>{
    const mid=normalizeId(
      r['Número de matrícula'] ??
      r['Numero de matrícula'] ??
      r['Numero de matricula'] ??
      r['Matrícula'] ??
      r['Matricula']
    );
    if(!mid)return;

    inscritosSourceIds.add(mid);

    const respondent=respondentById.get(mid);
    if(!respondent) return;

    const campusInscrito=r['Campus']??r['Escuela: Campus: Nombre de la cuenta']??'';
    const programaInscrito=
      r['Programa de interés'] ??
      r['Programa de Interés'] ??
      r['Programa de interes'] ??
      r['Programa de Interes'] ??
      '';

    const campus=campusInscrito
      ? displayCampus(campusInscrito)
      : respondent.campus;

    const campusNorm=campusInscrito
      ? normalizeCampus(campusInscrito)
      : respondent.campus_norm;

    const existing=inscritosMatchedById.get(mid);

    // Si hay filas duplicadas para una matrícula, conservamos el primer match,
    // pero completamos Campus/Programa si la primera fila venía vacía.
    if(!existing){
      inscritosMatchedById.set(mid,{
        campus,
        campus_norm:campusNorm,
        campus_inscrito:cleanText(campusInscrito)||campus,
        programa_inscrito:cleanText(programaInscrito)
      });
    }else{
      if(!existing.campus_inscrito && campusInscrito){
        existing.campus_inscrito=cleanText(campusInscrito);
      }
      if(!existing.programa_inscrito && programaInscrito){
        existing.programa_inscrito=cleanText(programaInscrito);
      }
    }
  });

  // =====================================================
  // ESTADO ACTUAL DE LOS RESPONDIENTES
  // La encuesta (Alumnos) sigue siendo el universo maestro.
  //
  // Para el PRONÓSTICO:
  //   Estatus = Prospecto  -> universo pendiente a pronosticar
  //   Estatus = Inscrito   -> continuidad ya observada
  //
  // Para la VALIDACIÓN provisional:
  //   Inscrito = 1
  //   cualquier otro estatus = 0
  // =====================================================
  const prospectRespondentById=new Map();
  const enrolledRespondentById=new Map();
  const statusMatchedRespondentById=new Map();

  respondentById.forEach((respondent,mid)=>{
    const info=vivaById.get(mid);
    if(!info) return;

    statusMatchedRespondentById.set(mid,info);

    if(isEnrolledStatus(info.estatus)){
      enrolledRespondentById.set(mid,info);
    }else if(isProspectStatus(info.estatus)){
      prospectRespondentById.set(mid,info);
    }
  });

  const baseStatusEnrolledTotal=[...vivaById.values()].reduce(
    (n,info)=>n+(isEnrolledStatus(info.estatus)?1:0),0
  );
  const baseStatusProspectTotal=[...vivaById.values()].reduce(
    (n,info)=>n+(isProspectStatus(info.estatus)?1:0),0
  );

  const enrolledNoSurvey=Math.max(
    0,
    baseStatusEnrolledTotal-enrolledRespondentById.size
  );

  // =====================================================
  // VALIDACIÓN PARCIAL DEL MODELO
  //
  // Universo = todos los respondientes de Alumnos.
  // Resultado observado actual:
  //   Estatus Inscrito = 1
  //   cualquier otro Estatus = 0
  //
  // OJO: es una fotografía provisional; varios No inscritos
  // todavía pueden convertirse antes del cierre.
  // =====================================================

  const validationRespondentIdsByCampus=new Map();

  respondentById.forEach((respondent,mid)=>{
    const vivaInfo=vivaById.get(mid);
    const inscritoInfo=inscritosMatchedById.get(mid);

    const campusNorm=
      vivaInfo?.campus_norm ||
      inscritoInfo?.campus_norm ||
      respondent.campus_norm;

    if(!validationRespondentIdsByCampus.has(campusNorm)){
      validationRespondentIdsByCampus.set(campusNorm,new Set());
    }
    validationRespondentIdsByCampus.get(campusNorm).add(mid);
  });

  function validationSummary(ids){
    const mids=[...(ids||[])];

    const propRows=PROP_ORDER.map(label=>({
      label,
      respondents:0,
      enrolled:0,
      prospect:0,
      other_not_enrolled:0,
      predicted_sum:0,
      predicted_n:0
    }));

    const propMap=new Map(propRows.map(r=>[r.label,r]));

    const priorityRows=[1,2,3,4,5].map(value=>({
      value,
      respondents:0,
      enrolled:0,
      predicted_sum:0,
      predicted_n:0,
      propensity_counts:Object.fromEntries(PROP_ORDER.map(label=>[label,0]))
    }));
    const priorityMap=new Map(priorityRows.map(r=>[r.value,r]));

    let enrolled=0;
    let prospect=0;
    let otherNotEnrolled=0;
    let enrolledHigh=0;
    let respondentProbSum=0,respondentProbN=0;
    let enrolledProbSum=0,enrolledProbN=0;

    mids.forEach(mid=>{
      const respondent=respondentById.get(mid);
      if(!respondent) return;

      const row=respondent.row;
      const prop=cleanText(row['Propension_Relativa']);
      const prob=asFloat(row['Prob_Continuidad_2027_Calibrada']);
      const priority=asInt(row['prioridad_tecmilenio']);

      const currentInfo=vivaById.get(mid);
      const isEnrolled=!!(currentInfo && isEnrolledStatus(currentInfo.estatus));
      const isProspect=!!(currentInfo && isProspectStatus(currentInfo.estatus));

      if(prob!==null){
        respondentProbSum+=prob;
        respondentProbN++;
      }

      if(isEnrolled){
        enrolled++;
        if(prob!==null){
          enrolledProbSum+=prob;
          enrolledProbN++;
        }
        if(['Alta','Muy alta'].includes(prop)) enrolledHigh++;
      }else if(isProspect){
        prospect++;
      }else{
        otherNotEnrolled++;
      }

      const pr=priorityMap.get(priority);
      if(pr){
        pr.respondents++;
        if(isEnrolled) pr.enrolled++;
        if(prob!==null){
          pr.predicted_sum+=prob;
          pr.predicted_n++;
        }
        if(prop in pr.propensity_counts){
          pr.propensity_counts[prop]++;
        }
      }

      const p=propMap.get(prop);
      if(p){
        p.respondents++;
        if(isEnrolled) p.enrolled++;
        else if(isProspect) p.prospect++;
        else p.other_not_enrolled++;

        if(prob!==null){
          p.predicted_sum+=prob;
          p.predicted_n++;
        }
      }
    });

    return {
      respondent_total:mids.length,
      enrolled,
      prospect,
      other_not_enrolled:otherNotEnrolled,
      observed_conversion:mids.length?enrolled/mids.length:0,
      enrolled_high_share:enrolled?enrolledHigh/enrolled:0,
      avg_predicted_respondents:respondentProbN?respondentProbSum/respondentProbN:null,
      avg_predicted_enrolled:enrolledProbN?enrolledProbSum/enrolledProbN:null,

      propensity:propRows.map(r=>({
        label:r.label,
        respondents:r.respondents,
        enrolled:r.enrolled,
        prospect:r.prospect,
        other_not_enrolled:r.other_not_enrolled,
        observed_conversion:r.respondents?r.enrolled/r.respondents:0,
        avg_predicted:r.predicted_n?r.predicted_sum/r.predicted_n:null
      })),

      priority:priorityRows.map(r=>({
        value:r.value,
        respondents:r.respondents,
        enrolled:r.enrolled,
        observed_conversion:r.respondents?r.enrolled/r.respondents:0,
        avg_predicted:r.predicted_n?r.predicted_sum/r.predicted_n:null,
        propensity_counts:r.propensity_counts,
        propensity_pcts:Object.fromEntries(
          PROP_ORDER.map(label=>[label,r.respondents?r.propensity_counts[label]/r.respondents:0])
        )
      }))
    };
  }

  const validationByCampus={};
  validationRespondentIdsByCampus.forEach((ids,norm)=>{
    validationByCampus[norm]=validationSummary(ids);
  });

  const validationOverall=
    validationSummary(new Set(respondentById.keys()));

  // =====================================================
  // CAMPUS OPERATIVO DE TODOS LOS RESPONDIENTES
  // Se usa para el explorador completo y para el denominador
  // de la tasa proyectada 2027 por campus.
  // Prioridad de fuente: Base Viva > Inscritos > Encuesta.
  // =====================================================
  const respondentIdsByCampus=new Map();
  respondentById.forEach((info,mid)=>{
    const vivaInfo=vivaById.get(mid);
    const inscritoInfo=inscritosMatchedById.get(mid);

    const campusNorm=
      vivaInfo?.campus_norm ||
      inscritoInfo?.campus_norm ||
      info.campus_norm;

    if(!respondentIdsByCampus.has(campusNorm)){
      respondentIdsByCampus.set(campusNorm,new Set());
    }
    respondentIdsByCampus.get(campusNorm).add(mid);
  });

  // =====================================================
  // 4) RESPONDIENTES CON ESTATUS = PROSPECTO
  // Único universo pendiente usado por el pronóstico.
  // =====================================================
  const active=[];
  alumnosAll.forEach(r=>{
    const mid=normalizeId(r['Matrícula']??r['Matricula']);

    if(mid&&prospectRespondentById.has(mid)){
      const info=prospectRespondentById.get(mid);

      active.push({
        ...r,
        _matricula:mid,
        _campus:info.campus||displayCampus(r['Campus']),
        _campus_norm:info.campus_norm||normalizeCampus(r['Campus'])
      });
    }
  });

  const histByCampus=groupBy(historico,r=>normalizeCampus(r['Campus']));

  const baseIdsByCampus=new Map();
  prospectRespondentById.forEach((info,mid)=>{
    if(!baseIdsByCampus.has(info.campus_norm))baseIdsByCampus.set(info.campus_norm,new Set());
    baseIdsByCampus.get(info.campus_norm).add(mid);
  });

  const inscritosIdsByCampus=new Map();
  enrolledRespondentById.forEach((info,mid)=>{
    if(!inscritosIdsByCampus.has(info.campus_norm))inscritosIdsByCampus.set(info.campus_norm,new Set());
    inscritosIdsByCampus.get(info.campus_norm).add(mid);
  });

  const activeByCampus=groupBy(active,r=>r._campus_norm);

  const campusSet=new Set([
    ...histByCampus.keys(),
    ...baseIdsByCampus.keys(),
    ...activeByCampus.keys(),
    ...inscritosIdsByCampus.keys(),
    ...respondentIdsByCampus.keys(),
  ]);

  const campusNorms=[...campusSet]
    .filter(Boolean)
    .sort((a,b)=>displayCampus(a).localeCompare(displayCampus(b),'es'));

  function summarizeGroup(norm,rows,histRows,baseCount,inscritosCount,respondentCount){
    const probs=rows.map(r=>asFloat(r['Prob_Continuidad_2027_Calibrada'])).filter(v=>v!==null).map(p=>Math.min(1,Math.max(0,p)));
    const sumProb=probs.reduce((a,b)=>a+b,0);
    const modelMean=probs.length?sumProb/probs.length:null;
    let modelLo=null,modelHi=null;
    if(probs.length){ const variance=probs.reduce((a,p)=>a+p*(1-p),0),sd=Math.sqrt(variance); modelLo=Math.max(0,(sumProb-1.96*sd)/probs.length); modelHi=Math.min(1,(sumProb+1.96*sd)/probs.length); }
    const hist3=histRows.length?weightedHist(histRows,[2024,2025,2026]):null;
    const hist2026=histRows.length?weightedHist(histRows,[2026]):null;

    // La base operativa son exclusivamente respondientes
    // cuyo Estatus actual es Prospecto.
    const operationalBase=baseCount||rows.length;
    const baseMethod=baseCount?'Respondientes con Estatus Prospecto':'prospectos con encuesta/modelo';
    let refRate=null,lowRate=null,highRate=null,rateMethod='sin información suficiente';
    if(hist3!==null&&modelHi!==null){ refRate=(hist3+modelHi)/2; lowRate=Math.min(hist3,modelMean??hist3); highRate=Math.max(hist3,modelHi); rateMethod='punto medio entre promedio 3A ponderado y límite superior del modelo'; }
    else if(modelMean!==null){ refRate=modelMean;lowRate=modelLo??modelMean;highRate=modelHi??modelMean;rateMethod='modelo calibrado (sin histórico comparable)'; }
    else if(hist3!==null){ refRate=lowRate=highRate=hist3;rateMethod='promedio histórico 3A ponderado (sin modelo comparable)'; }
    const cont=refRate===null?null:Math.round(operationalBase*refRate);
    const contLo=lowRate===null?null:Math.round(operationalBase*lowRate);
    const contHi=highRate===null?null:Math.round(operationalBase*highRate);

    // Pronóstico total = pronóstico entre Prospectos respondientes
    //                  + respondientes con Estatus Inscrito.
    const inscritosConfirmados=Number(inscritosCount||0);
    const forecastTotal=cont===null?inscritosConfirmados:cont+inscritosConfirmados;
    const forecastTotalLow=contLo===null?inscritosConfirmados:contLo+inscritosConfirmados;
    const forecastTotalHigh=contHi===null?inscritosConfirmados:contHi+inscritosConfirmados;

    // Para la gráfica de tendencia 2027:
    // tasa proyectada = pronóstico total / total de respondientes.
    const respondentTotal=Number(respondentCount||0);
    const trendRate2027=respondentTotal?forecastTotal/respondentTotal:null;
    const trendRateLow2027=respondentTotal?forecastTotalLow/respondentTotal:null;
    const trendRateHigh2027=respondentTotal?forecastTotalHigh/respondentTotal:null;

    const propCounts=Object.fromEntries(PROP_ORDER.map(k=>[k,0]));
    rows.forEach(r=>{const v=cleanText(r['Propension_Relativa']);if(v in propCounts)propCounts[v]++;});
    const prop=PROP_ORDER.map(k=>({label:k,count:propCounts[k],pct:rows.length?propCounts[k]/rows.length:0}));
    const highCount=propCounts['Alta']+propCounts['Muy alta'],lowCount=propCounts['Baja']+propCounts['Muy baja'];
    const highPct=rows.length?highCount/rows.length:0,lowPct=rows.length?lowCount/rows.length:0;

    const clMap=new Map(); rows.forEach(r=>{const k=cleanText(r['Cluster_Nombre'])||'Sin cluster';clMap.set(k,(clMap.get(k)||0)+1);});
    const clusters=[...clMap.entries()].sort((a,b)=>b[1]-a[1]).map(([label,count])=>({label,count,pct:rows.length?count/rows.length:0}));
    const factors=Object.entries(FACTOR_LABELS).map(([col,label])=>({label,count:countBinary(rows,col)})).filter(x=>x.count).map(x=>({...x,pct:rows.length?x.count/rows.length:0})).sort((a,b)=>b.count-a.count);
    const reasons=Object.entries(REASON_LABELS).map(([col,label])=>({label,count:countBinary(rows,col)})).filter(x=>x.count).map(x=>({...x,pct:rows.length?x.count/rows.length:0})).sort((a,b)=>b.count-a.count);
    const oppCount=rows.reduce((n,r)=>{ const level=cleanText(r['Prioridad_Convencimiento']); const flag=cleanText(r['Oportunidad_Convencimiento']).toLowerCase(); return n+(OPPORTUNITY_LEVELS.has(level)||flag==='si'||flag==='sí'?1:0); },0);
    const offerCount=rows.reduce((n,r)=>n+(asFloat(r['programa_ofertado'])===1?1:0),0);
    const delta=refRate!==null&&hist2026!==null?refRate-hist2026:null;
    let status;
    if(hist2026===null) status='Sin histórico comparable';
    else if(delta<=-0.03&&lowPct>=highPct+0.05) status='Riesgo de caída';
    else if(delta>=0.03&&highPct>=lowPct+0.05) status='Potencial de mejora';
    else if(Math.abs(delta)<=0.03) status='Estable';
    else status='Seguimiento';
    let explanation;
    if(status==='Riesgo de caída') explanation=`Se proyectan ${fmt.format(cont||0)} continuadores. La referencia queda ${(Math.abs(delta)*100).toFixed(1)} pp por debajo de 2026 y ${fmt.format(lowCount)} alumnos (${(lowPct*100).toFixed(0)}%) están en propensión baja o muy baja, frente a ${fmt.format(highCount)} (${(highPct*100).toFixed(0)}%) en alta o muy alta.`;
    else if(status==='Potencial de mejora') explanation=`Se proyectan ${fmt.format(cont||0)} continuadores. La referencia supera 2026 en ${(delta*100).toFixed(1)} pp y ${fmt.format(highCount)} alumnos (${(highPct*100).toFixed(0)}%) están en propensión alta o muy alta, frente a ${fmt.format(lowCount)} (${(lowPct*100).toFixed(0)}%) en baja o muy baja.`;
    else if(status==='Estable') explanation=`Se proyectan ${fmt.format(cont||0)} continuadores y la referencia se mantiene cerca de 2026 (${delta>=0?'+':''}${(delta*100).toFixed(1)} pp). La cohorte activa tiene ${fmt.format(highCount)} alumnos en propensión alta/muy alta y ${fmt.format(lowCount)} en baja/muy baja.`;
    else if(status==='Seguimiento') explanation=`Se proyectan ${fmt.format(cont||0)} continuadores. La tasa de referencia y la composición de propensión envían señales mixtas: ${fmt.format(highCount)} alumnos en alta/muy alta y ${fmt.format(lowCount)} en baja/muy baja.`;
    else explanation=`Se proyectan ${fmt.format(cont||0)} continuadores con la señal disponible del modelo. No existe histórico comparable suficiente para este campus.`;
    explanation+=` A este pronóstico de Prospectos se suman ${fmt.format(inscritosConfirmados)} respondientes con Estatus Inscrito, para un pronóstico total de ${fmt.format(forecastTotal)}.`;
    if(oppCount) explanation+=` Además, ${fmt.format(oppCount)} respondentes activos aparecen como oportunidad de convencimiento media-alta, alta o muy alta.`;

    const histAbs=histRows.length?historySummary(histRows):Array.from({length:6},(_,i)=>({year:2021+i,n:0,rate:null,continuers:0}));
    histAbs.push({year:2027,n:respondentTotal,rate:trendRate2027,continuers:forecastTotal||0,forecast:true});
    return {
      campus_norm:norm,
      campus:norm==='__ALL__'?'Todos los campus':displayCampus(norm),
      survey_n:rows.length,
      base_viva_n:Number(baseCount||0),
      inscritos_confirmados:inscritosConfirmados,
      respondent_total_n:respondentTotal,
      trend_rate_2027:trendRate2027,
      trend_rate_low_2027:trendRateLow2027,
      trend_rate_high_2027:trendRateHigh2027,
      operational_base:Number(operationalBase||0),
      operational_base_method:baseMethod,
      continuers_expected:cont,
      continuers_low:contLo,
      continuers_high:contHi,
      forecast_total:forecastTotal,
      forecast_total_low:forecastTotalLow,
      forecast_total_high:forecastTotalHigh,
      forecast_rate:refRate,forecast_rate_low:lowRate,forecast_rate_high:highRate,model_mean:modelMean,model_upper:modelHi,hist3_weighted:hist3,hist2026,state:status,delta_vs_2026_pp:delta===null?null:delta*100,rate_method:rateMethod,explanation,
      history:histAbs,propensity:prop,clusters,factors,reasons,opportunity_count:oppCount,program_offered_count:offerCount,program_offered_pct:rows.length?offerCount/rows.length:0,nps:average(rows,'NPS'),orgullo:average(rows,'orgullo'),seguridad:average(rows,'seguridad_campus'),infra:average(rows,'Satisfaccion infra')
    };
  }

  const summaries={};
  campusNorms.forEach(norm=>summaries[norm]=summarizeGroup(
    norm,
    activeByCampus.get(norm)||[],
    histByCampus.get(norm)||[],
    baseIdsByCampus.get(norm)?.size||0,
    inscritosIdsByCampus.get(norm)?.size||0,
    respondentIdsByCampus.get(norm)?.size||0
  ));

  const overall=summarizeGroup(
    '__ALL__',
    active,
    historico,
    vivaById.size,
    inscritosMatchedById.size,
    respondentById.size
  );

  overall.operational_base=campusNorms.reduce((a,n)=>a+(summaries[n].operational_base||0),0);
  overall.continuers_expected=campusNorms.reduce((a,n)=>a+(summaries[n].continuers_expected||0),0);
  overall.continuers_low=campusNorms.reduce((a,n)=>a+(summaries[n].continuers_low||0),0);
  overall.continuers_high=campusNorms.reduce((a,n)=>a+(summaries[n].continuers_high||0),0);
  overall.inscritos_confirmados=inscritosMatchedById.size;
  overall.forecast_total=overall.continuers_expected+overall.inscritos_confirmados;
  overall.forecast_total_low=overall.continuers_low+overall.inscritos_confirmados;
  overall.forecast_total_high=overall.continuers_high+overall.inscritos_confirmados;
  overall.respondent_total_n=respondentById.size;
  overall.trend_rate_2027=overall.respondent_total_n?overall.forecast_total/overall.respondent_total_n:null;
  overall.trend_rate_low_2027=overall.respondent_total_n?overall.forecast_total_low/overall.respondent_total_n:null;
  overall.trend_rate_high_2027=overall.respondent_total_n?overall.forecast_total_high/overall.respondent_total_n:null;
  if(overall.operational_base) overall.forecast_rate=overall.continuers_expected/overall.operational_base;
  if(overall.history?.length){
    const last=overall.history[overall.history.length-1];
    last.n=overall.respondent_total_n;
    last.continuers=overall.forecast_total;
    last.rate=overall.trend_rate_2027;
  }
  const highCount=(overall.propensity||[]).filter(x=>['Alta','Muy alta'].includes(x.label)).reduce((a,x)=>a+x.count,0);
  const lowCount=(overall.propensity||[]).filter(x=>['Baja','Muy baja'].includes(x.label)).reduce((a,x)=>a+x.count,0);
  const highPct=overall.survey_n?highCount/overall.survey_n:0,lowPct=overall.survey_n?lowCount/overall.survey_n:0;
  overall.explanation=`Sobre la Base Viva actual de ${fmt.format(overall.operational_base)} alumnos se proyectan ${fmt.format(overall.continuers_expected)} continuadores. Además, ${fmt.format(overall.inscritos_confirmados)} respondientes hacen match con INSCRITOS_2027 y ya están confirmados. El pronóstico total es ${fmt.format(overall.forecast_total)}. Entre los respondentes que siguen en Base Viva, ${fmt.format(highCount)} (${(highPct*100).toFixed(0)}%) están en propensión alta/muy alta y ${fmt.format(lowCount)} (${(lowPct*100).toFixed(0)}%) en baja/muy baja. Se identifican ${fmt.format(overall.opportunity_count)} oportunidades de convencimiento media-alta, alta o muy alta.`;
  summaries['__ALL__']=overall;

  // =====================================================
  // EXPLORADOR: CONSERVAR TODOS LOS RESPONDIENTES
  // Ya no se eliminan quienes no hacen match con Base Viva.
  // Cada matrícula recibe dos etiquetas independientes:
  // - En Base Viva: Sí / No
  // - Inscrito: Sí / No
  // =====================================================
  const students=[...respondentById.entries()].map(([mid,respondent])=>{
    const r=respondent.row;
    const vivaInfo=vivaById.get(mid);
    const statusInfo=vivaById.get(mid);
    const inscritoInfo=inscritosMatchedById.get(mid);

    const campus=
      vivaInfo?.campus ||
      inscritoInfo?.campus ||
      respondent.campus;

    const campusNorm=
      vivaInfo?.campus_norm ||
      inscritoInfo?.campus_norm ||
      respondent.campus_norm;

    const factors=Object.entries(FACTOR_LABELS)
      .filter(([col])=>asFloat(r[col])===1)
      .map(([,label])=>label);

    const reasons=Object.entries(REASON_LABELS)
      .filter(([col])=>asFloat(r[col])===1)
      .map(([,label])=>label);

    return {
      matricula:mid,
      campus,
      campus_norm:campusNorm,
      en_base_viva:!!vivaInfo,
      en_base_estatus:!!statusInfo,
      estatus_base_viva:statusInfo?.estatus||'',
      es_prospecto:!!(statusInfo&&isProspectStatus(statusInfo.estatus)),
      inscrito:!!(statusInfo&&isEnrolledStatus(statusInfo.estatus)),
      campus_inscrito:inscritoInfo?.campus_inscrito||'',
      programa_inscrito:inscritoInfo?.programa_inscrito||'',
      prob:asFloat(r['Prob_Continuidad_2027_Calibrada']),
      propension:cleanText(r['Propension_Relativa']),
      percentil_global:asFloat(r['Percentil_Propension_Global']??r['Percentil_Propension']),
      percentil_campus:asFloat(r['Percentil_Propension_Campus']),
      potencial:cleanText(r['Potencial_Estructural']),
      percentil_estructural:asFloat(r['Percentil_Estructural']),
      brecha_oportunidad:asFloat(r['Brecha_Oportunidad']),
      prioridad_convencimiento:cleanText(r['Prioridad_Convencimiento']),
      oportunidad:cleanText(r['Oportunidad_Convencimiento']),
      lectura_oportunidad:cleanText(r['Lectura_Oportunidad']),
      resumen_individual:cleanText(r['Resumen_Individual']),
      cluster_id:asInt(r['Cluster_ID']),
      cluster_nombre:cleanText(r['Cluster_Nombre']),
      cluster_descripcion:cleanText(r['Cluster_Descripcion']),
      cluster_palanca:cleanText(r['Cluster_Palanca']),
      carrera:cleanText(r['¿Qué carrera te gustaría estudiar?']),
      planea_tecmilenio:cleanText(r['¿Planeas iniciar tus estudios profesionales en Tecmilenio?']),
      universidad:cleanText(r['¿En dónde planeas realizar tus estudios profesionales?']),
      programa_interes:cleanText(r['De la oferta institucional de carreras que ofrece Tecmilenio en todos sus campus a nivel nacional, ¿hay alguna que te interese? (Nota: posteriormente, acércate con tu asesor de admisiones para conoc']),
      modalidad:cleanText(r['modalidad_clases']),
      movilidad:cleanText(r['movilidad_ciudad']),
      prioridad:asFloat(r['prioridad_tecmilenio']),
      campus_interes:cleanText(r['¿En qué campus de Tecmilenio te gustaría estudiar tu carrera profesional?']),
      orientacion:cleanText(r['¿Te gustaría recibir una sesión de orientación vocacional personalizada?']),
      seguridad_carrera:asFloat(r['seguridad_carrera']),
      nps:asFloat(r['NPS']),
      orgullo:asFloat(r['orgullo']),
      seguridad_campus:asFloat(r['seguridad_campus']),
      infra:asFloat(r['Satisfaccion infra']),
      beca:asFloat(r['Beca']),
      programa_ofertado:asFloat(r['programa_ofertado']),
      factores:factors,
      barreras:reasons
    };
  });


  const baseEstatusRespondentMatch=[...respondentById.keys()].reduce(
    (n,mid)=>n+(vivaById.has(mid)?1:0),0
  );
  const baseEstatusNoSurvey=Math.max(
    0,
    vivaById.size-baseEstatusRespondentMatch
  );
  const respondentNoBaseEstatus=Math.max(
    0,
    respondentById.size-baseEstatusRespondentMatch
  );

  const baseVivaStatusCounts={};
  vivaById.forEach(info=>{
    const status=cleanText(info.estatus)||'Sin estatus';
    baseVivaStatusCounts[status]=(baseVivaStatusCounts[status]||0)+1;
  });


  return {
    meta:{
      survey_total:alumnosAll.length,
      respondent_unique_total:respondentById.size,
      base_viva_total:vivaById.size,
      base_estatus_total:vivaById.size,
      base_estatus_respondent_match:baseEstatusRespondentMatch,
      base_estatus_no_survey:baseEstatusNoSurvey,
      prospectos_base_total:baseStatusProspectTotal,
      inscritos_base_total:baseStatusEnrolledTotal,
      prospectos_respondientes:prospectRespondentById.size,
      inscritos_respondientes:enrolledRespondentById.size,
      respondent_no_base_estatus:respondentNoBaseEstatus,
      base_viva_status_counts:baseVivaStatusCounts,
      matched_active:active.length,
      coverage_active:respondentById.size?active.length/respondentById.size:0,
      inscritos_source_unique_total:inscritosSourceIds.size,
      inscritos_respondent_match_total:enrolledRespondentById.size,
      inscritos_program_match_total:inscritosMatchedById.size,
      inscritos_no_survey_total:enrolledNoSurvey,
      respondents_outside_base_viva:students.filter(s=>!s.en_base_viva).length,
      respondents_not_enrolled:students.filter(s=>!s.inscrito).length,
      analysis_file:'base_continuidad_2027_dashboard.xlsx',
      base_viva_file:'BASE_VIVA.xlsx',
      inscritos_file:'INSCRITOS_2027.xlsx',
      loaded_at:new Date().toISOString(),
      method_note:'La encuesta sigue siendo el universo maestro del modelo. BASE_VIVA.xlsx aporta el Estatus actual. El pronóstico pendiente se calcula únicamente entre respondientes con Estatus Prospecto y se suman los respondientes con Estatus Inscrito. En la validación provisional, Inscrito=1 y cualquier otro Estatus=0; el análisis de Prioridad 1–5 se conserva como diagnóstico.'
    },
    campuses:campusNorms.map(n=>summaries[n]),
    overall:summaries['__ALL__'],
    students,
    validation:{overall:validationOverall,by_campus:validationByCampus},
    clusters_catalog:catalogo
  };
}



function currentSummary(){
  if(state.selectedCampus==='__ALL__') return state.data.overall;
  return state.data.campuses.find(c=>c.campus_norm===state.selectedCampus) || state.data.overall;
}
function studentsForCampus(){
  const all=state.data.students;
  return state.selectedCampus==='__ALL__'?all:all.filter(s=>s.campus_norm===state.selectedCampus);
}
function campusOptions(){
  $('campusSelect').innerHTML = `<option value="__ALL__">Todos los campus</option>` + state.data.campuses.map(c=>`<option value="${safe(c.campus_norm)}">${safe(c.campus)}</option>`).join('');
  $('campusSelect').value=state.selectedCampus;
}

function renderKPIs(s){
  $('kpiBase').textContent=fmt.format(s.base_viva_n||0);

  $('kpiSurvey').textContent=fmt.format(s.respondent_total_n||0);
  const coverage=s.respondent_total_n? s.base_viva_n/s.respondent_total_n : 0;
  $('kpiCoverage').textContent=`${pct(coverage)} con Estatus Prospecto`;

  $('kpiBaseForecast').textContent=s.continuers_expected==null?'—':fmt.format(s.continuers_expected);
  $('kpiBaseForecastSub').textContent=`rango ${fmt.format(s.continuers_low||0)}–${fmt.format(s.continuers_high||0)}`;

  $('kpiInscritos').textContent=fmt.format(s.inscritos_confirmados||0);
  $('kpiInscritosSub').textContent='respondientes con Estatus Inscrito';

  $('kpiCont').textContent=s.forecast_total==null?'—':fmt.format(s.forecast_total);
  $('kpiContSub').textContent=`${fmt.format(s.continuers_expected||0)} esperados entre Prospectos + ${fmt.format(s.inscritos_confirmados||0)} ya inscritos`;

  $('kpiOpp').textContent=fmt.format(s.opportunity_count||0);

  const m=state.data.meta;
  $('sourceText').textContent=
    `${fmt.format(m.base_viva_total||0)} registros Base Viva · `+
    `${fmt.format(m.base_estatus_respondent_match||0)} matches con encuesta · `+
    `${fmt.format(m.prospectos_respondientes||0)} Prospectos · `+
    `${fmt.format(m.inscritos_respondientes||0)} Inscritos`;
}

function currentValidation(){
  if(!state.data?.validation) return null;

  if(state.selectedCampus==='__ALL__'){
    return state.data.validation.overall;
  }

  return state.data.validation.by_campus?.[state.selectedCampus] || {
    respondent_total:0,
    enrolled:0,
    prospect:0,
    other_not_enrolled:0,
    observed_conversion:0,
    enrolled_high_share:0,
    avg_predicted_respondents:null,
    avg_predicted_enrolled:null,
    propensity:PROP_ORDER.map(label=>({
      label,
      respondents:0,
      enrolled:0,
      prospect:0,
      other_not_enrolled:0,
      observed_conversion:0,
      avg_predicted:null
    })),
    priority:[1,2,3,4,5].map(value=>({
      value,
      respondents:0,
      enrolled:0,
      observed_conversion:0,
      avg_predicted:null,
      propensity_counts:Object.fromEntries(PROP_ORDER.map(label=>[label,0])),
      propensity_pcts:Object.fromEntries(PROP_ORDER.map(label=>[label,0]))
    }))
  };
}

function classificationAtThreshold(threshold){
  const rows=studentsForCampus();

  let tp=0,fp=0,tn=0,fn=0,missing=0,evaluated=0;

  rows.forEach(s=>{
    const hasStatus=!!cleanText(s.estatus_base_viva);
    const prob=asFloat(s.prob);

    if(!hasStatus || prob===null){
      missing++;
      return;
    }

    evaluated++;
    const actual=!!s.inscrito;
    const predicted=prob>=threshold;

    if(predicted && actual) tp++;
    else if(predicted && !actual) fp++;
    else if(!predicted && actual) fn++;
    else tn++;
  });

  const accuracy=evaluated?(tp+tn)/evaluated:null;
  const precision=(tp+fp)?tp/(tp+fp):null;
  const recall=(tp+fn)?tp/(tp+fn):null;
  const f1=(precision!==null && recall!==null && (precision+recall))
    ?2*precision*recall/(precision+recall)
    :null;

  return {threshold,tp,fp,tn,fn,missing,evaluated,accuracy,precision,recall,f1};
}

function renderValidation(){
  const v=currentValidation();
  if(!v) return;

  const m=state.data.meta||{};

  setText('coverageBaseTotal',fmt.format(m.base_estatus_total||0));
  setText('coverageBaseSurvey',fmt.format(m.base_estatus_respondent_match||0));
  setText('coverageBaseNoSurvey',fmt.format(m.base_estatus_no_survey||0));
  setText('coverageSurveyNoBase',fmt.format(m.respondent_no_base_estatus||0));
  setText('coverageEnrolledNoSurvey',fmt.format(m.inscritos_no_survey_total||0));

  setText('valRespondents',fmt.format(v.respondent_total||0));
  setText('valEnrolled',fmt.format(v.enrolled||0));
  setText('valConversion',pct(v.observed_conversion));
  setText('valProspect',fmt.format(v.prospect||0));
  setText('valOtherStatus',fmt.format(v.other_not_enrolled||0));
  setText('valHighShare',pct(v.enrolled_high_share));

  setText('valAvgRespondents',pct(v.avg_predicted_respondents));
  setText('valAvgEnrolled',pct(v.avg_predicted_enrolled));

  const c=classificationAtThreshold(state.validationThreshold);

  setText('valThreshold',state.validationThreshold.toFixed(2));
  setText('valEvaluated',fmt.format(c.evaluated||0));
  setText('valAccuracy',pct(c.accuracy));
  setText('valPrecision',pct(c.precision));
  setText('valRecall',pct(c.recall));
  setText('valF1',pct(c.f1));
  setText('valTP',fmt.format(c.tp||0));
  setText('valFP',fmt.format(c.fp||0));
  setText('valTN',fmt.format(c.tn||0));
  setText('valFN',fmt.format(c.fn||0));
  setText('valMissingStatus',fmt.format(c.missing||0));

  const rows=v.propensity||[];

  $('validationTableBody').innerHTML=rows.map(r=>`
    <tr>
      <td><strong>${safe(r.label)}</strong></td>
      <td class="num">${fmt.format(r.respondents||0)}</td>
      <td class="num">${fmt.format(r.enrolled||0)}</td>
      <td class="num"><strong>${pct(r.observed_conversion)}</strong></td>
      <td class="num">${fmt.format(r.prospect||0)}</td>
      <td class="num">${fmt.format(r.other_not_enrolled||0)}</td>
      <td class="num">${pct(r.avg_predicted)}</td>
    </tr>
  `).join('');

  const priorityRows=v.priority||[];

  $('priorityConversionTableBody').innerHTML=priorityRows.map(r=>`
    <tr>
      <td><strong>${r.value}</strong></td>
      <td class="num">${fmt.format(r.respondents||0)}</td>
      <td class="num">${fmt.format(r.enrolled||0)}</td>
      <td class="num"><strong>${pct(r.observed_conversion)}</strong></td>
      <td class="num">${pct(r.avg_predicted)}</td>
    </tr>
  `).join('');

  $('priorityPropensityTableBody').innerHTML=priorityRows.map(r=>`
    <tr>
      <td><strong>Prioridad ${r.value}</strong><br><small>${fmt.format(r.respondents||0)} alumnos</small></td>
      ${PROP_ORDER.map(label=>{
        const count=r.propensity_counts?.[label]||0;
        const share=r.propensity_pcts?.[label]||0;
        return `<td class="num">${fmt.format(count)}<br><small>${pct(share)}</small></td>`;
      }).join('')}
    </tr>
  `).join('');

  const campusText=
    state.selectedCampus==='__ALL__'
      ? 'todos los campus'
      : `el campus ${safe(currentSummary().campus)}`;

  $('validationNote').innerHTML=
    `Validación parcial sobre <strong>${fmt.format(v.respondent_total||0)} respondientes</strong> de ${campusText}. `+
    `El resultado observado actual es <strong>Estatus Inscrito = 1</strong> y cualquier otro Estatus = 0. `+
    `Con umbral <strong>${state.validationThreshold.toFixed(2)}</strong>, el modelo identifica ${fmt.format(c.tp)} inscritos correctamente y deja ${fmt.format(c.fn)} inscritos por debajo del corte. `+
    `Estas métricas son provisionales porque varios alumnos No inscritos todavía pueden convertir antes del cierre.`;
}


function setDashboardTab(tab){
  state.selectedDashboardTab=tab==='validation'?'validation':'forecast';

  const isValidation=state.selectedDashboardTab==='validation';

  $('forecastTab').classList.toggle('active',!isValidation);
  $('validationTab').classList.toggle('active',isValidation);

  $('tabForecast').classList.toggle('active',!isValidation);
  $('tabValidation').classList.toggle('active',isValidation);

  $('tabForecast').setAttribute('aria-selected',String(!isValidation));
  $('tabValidation').setAttribute('aria-selected',String(isValidation));

  if(isValidation){
    renderValidation();
  }
}


function renderDataNotes(){
  const m=state.data.meta;
  $('surveyNote').innerHTML=
    `La encuesta sigue siendo el universo maestro con <strong>${fmt.format(m.respondent_unique_total)} respondientes únicos</strong>. `+
    `La Base Viva actual contiene <strong>${fmt.format(m.base_viva_total)}</strong> matrículas y ${fmt.format(m.base_estatus_respondent_match)} hacen match con la encuesta; `+
    `${fmt.format(m.base_estatus_no_survey)} pertenecen a alumnos sin encuesta. Entre los respondientes hay `+
    `<strong>${fmt.format(m.prospectos_respondientes)} Prospectos</strong> y <strong>${fmt.format(m.inscritos_respondientes)} Inscritos</strong>.`;

  $('methodSurvey').textContent=
    `Pronóstico total = pronóstico sobre los respondientes con Estatus Prospecto + respondientes con Estatus Inscrito. `+
    `La tendencia 2027 usa pronóstico total / respondientes del dashboard. La evaluación provisional toma Estatus Inscrito como 1 y cualquier otro Estatus como 0.`;
}

function renderHistory(s){
  $('stateBadge').textContent=s.state;
  $('stateBadge').className=`state ${cssState(s.state)}`;
  const rateLo=s.trend_rate_low_2027;
  const rateHi=s.trend_rate_high_2027;
  $('rangeText').textContent=(rateLo==null||rateHi==null)?'—':`Rango tasa ${pct(rateLo)}–${pct(rateHi)}`;

  const data=s.history||[];
  const width=720,height=220,pad={l:48,r:14,t:15,b:34};
  const validRates=data.map(d=>d.rate).filter(v=>v!==null&&v!==undefined&&Number.isFinite(Number(v))).map(Number);
  const observedMax=validRates.length?Math.max(...validRates):0.30;
  // Eje porcentual desde 0, con techo redondeado a decenas y mínimo de 40%.
  const maxRate=Math.max(0.40,Math.ceil(observedMax*10)/10);
  const x=i=>pad.l+(width-pad.l-pad.r)*(i/(Math.max(1,data.length-1)));
  const y=v=>height-pad.b-(height-pad.t-pad.b)*(Number(v)/maxRate);

  let grid='';
  for(let i=0;i<=4;i++){
    const val=maxRate*i/4;
    const yy=y(val);
    grid+=`<line x1="${pad.l}" y1="${yy}" x2="${width-pad.r}" y2="${yy}" stroke="#e9efec"/><text x="${pad.l-7}" y="${yy+3}" text-anchor="end" font-size="9" fill="#71817b">${Math.round(val*100)}%</text>`;
  }

  const histPoints=data.map((d,i)=>({d,i})).filter(({d})=>!d.forecast&&d.rate!=null&&Number.isFinite(Number(d.rate)));
  const histPolyline=histPoints.length>1?`<polyline points="${histPoints.map(({d,i})=>`${x(i)},${y(d.rate)}`).join(' ')}" fill="none" stroke="#667a72" stroke-width="2.3"/>`:'';
  const forecastItem=data.map((d,i)=>({d,i})).find(({d})=>d.forecast&&d.rate!=null&&Number.isFinite(Number(d.rate)));
  const lastHist=histPoints.length?histPoints[histPoints.length-1]:null;
  const forecastLine=(forecastItem&&lastHist)?`<line x1="${x(lastHist.i)}" y1="${y(lastHist.d.rate)}" x2="${x(forecastItem.i)}" y2="${y(forecastItem.d.rate)}" stroke="#16805d" stroke-width="3" stroke-dasharray="7 5"/>`:'';

  let marks='';
  data.forEach((d,i)=>{
    marks+=`<text x="${x(i)}" y="${height-12}" text-anchor="middle" font-size="9" fill="#667670">${d.year}</text>`;
    if(d.rate==null||!Number.isFinite(Number(d.rate))) return;
    const forecast=d.forecast;
    marks+=`<circle cx="${x(i)}" cy="${y(d.rate)}" r="${forecast?5:4}" fill="${forecast?'#15805d':'#6f7f79'}"><title>${d.year}: ${pct(d.rate)} continuidad · ${fmt.format(d.continuers||0)} continuadores · N ${fmt.format(d.n||0)}</title></circle>`;
  });

  $('historyChart').innerHTML=`<svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">${grid}${histPolyline}${forecastLine}${marks}</svg>`;
  $('historyTable').innerHTML=`<table class="history-mini"><thead><tr><th>Año</th>${data.map(d=>`<th>${d.year}</th>`).join('')}</tr></thead><tbody><tr><td>Base / respondientes 2027</td>${data.map(d=>`<td>${fmt.format(d.n||0)}</td>`).join('')}</tr><tr><td>Continuadores</td>${data.map(d=>`<td class="${d.forecast?'forecast':''}">${fmt.format(d.continuers||0)}</td>`).join('')}</tr><tr><td>Tasa</td>${data.map(d=>`<td>${d.rate==null?'—':pct(d.rate)}</td>`).join('')}</tr></tbody></table>`;
  $('forecastExplanation').textContent=s.explanation;
}

const clusterColors=['#1c8663','#3479ad','#98783e','#8b5e91','#718c56','#a9634f'];
function renderClusters(s){
  const rows=s.clusters||[]; const total=s.survey_n||0;
  let pos=0; const parts=[]; rows.forEach((r,i)=>{ const start=pos*100; pos+=r.pct; const end=pos*100; parts.push(`${clusterColors[i%clusterColors.length]} ${start}% ${end}%`); });
  $('clusterDonut').style.background=parts.length?`conic-gradient(${parts.join(',')})`:'#eef2f0';
  $('clusterTotal').textContent=fmt.format(total);
  $('clusterLegend').innerHTML=rows.slice(0,6).map((r,i)=>`<div class="legend-row"><span class="legend-dot" style="background:${clusterColors[i%clusterColors.length]}"></span><span>${safe(r.label)}</span><span class="sub"><strong>${fmt.format(r.count)}</strong> · ${pct(r.pct)}</span></div>`).join('')||'<p class="muted">Sin datos</p>';
  renderClusterDescriptions();
}

function renderClusterDescriptions(){
  const target=$('clusterDescriptions');
  if(!target) return;
  const rows=(state.data?.clusters_catalog||[]).slice().sort((a,b)=>(asInt(a['ID'])??99)-(asInt(b['ID'])??99));
  target.innerHTML=rows.length?rows.map((r,i)=>{
    const id=asInt(r['ID']);
    const nombre=cleanText(r['Nombre'])||`Cluster ${id??''}`;
    const desc=cleanText(r['Descripcion']);
    const palanca=cleanText(r['Palanca principal']);
    const mensaje=cleanText(r['Mensaje clave']).replace(/^"|"$/g,'');
    return `<div class="cluster-desc-card"><div class="cluster-desc-head"><span class="cluster-desc-id" style="background:${clusterColors[(id??i)%clusterColors.length]}">${id??i}</span><h4>${safe(nombre)}</h4></div><p>${safe(desc)}</p>${palanca?`<p class="lever">Palanca: ${safe(palanca)}</p>`:''}${mensaje?`<p>${safe(mensaje)}</p>`:''}</div>`;
  }).join(''):'<p class="muted">No hay catálogo de clusters disponible.</p>';
}

function propensityRowsForCampus(){
  const students=studentsForCampus();
  const counts=Object.fromEntries(PROP_ORDER.map(k=>[k,0]));
  students.forEach(s=>{
    if(s.propension in counts) counts[s.propension]++;
  });
  return PROP_ORDER.map(label=>({
    label,
    count:counts[label],
    pct:students.length?counts[label]/students.length:0
  }));
}

function setPropensityFilter(label,scrollToList=false){
  state.selectedPropensity=label||'all';
  state.selectedStudent=null;
  renderPropensity();
  renderStudentList();
  renderStudentDetail();

  if(scrollToList){
    document.querySelector('.student-explorer')?.scrollIntoView({
      behavior:'smooth',
      block:'start'
    });
  }
}

function renderPropensity(){
  const rows=propensityRowsForCampus();
  const max=Math.max(1,...rows.map(r=>r.count));

  $('propensityBars').innerHTML=rows.map(r=>`
    <button class="vbar propensity-bar-btn ${state.selectedPropensity===r.label?'active':''}"
            type="button"
            data-propensity="${safe(r.label)}"
            title="Ver alumnos con propensión ${safe(r.label)}">
      <span class="value">${fmt.format(r.count)}<br><small>${pct(r.pct)}</small></span>
      <span class="bar" style="height:${Math.max(3,r.count/max*115)}px"></span>
      <span class="label">${safe(r.label)}</span>
    </button>
  `).join('');

  $('propensityFilters').innerHTML=
    `<button type="button" class="prop-filter-chip ${state.selectedPropensity==='all'?'active':''}" data-prop-filter="all">Todas</button>`+
    rows.map(r=>`<button type="button" class="prop-filter-chip ${state.selectedPropensity===r.label?'active':''}" data-prop-filter="${safe(r.label)}">${safe(r.label)} · ${fmt.format(r.count)}</button>`).join('');

  document.querySelectorAll('[data-propensity]').forEach(el=>{
    el.addEventListener('click',()=>setPropensityFilter(el.dataset.propensity,true));
  });
  document.querySelectorAll('[data-prop-filter]').forEach(el=>{
    el.addEventListener('click',()=>setPropensityFilter(el.dataset.propFilter,true));
  });
}

function renderRank(id,rows){
  const top=(rows||[]).slice(0,6); const max=Math.max(1,...top.map(r=>r.count));
  $(id).innerHTML=top.length?top.map(r=>`<div class="rank-row"><span class="rank-label" title="${safe(r.label)}">${safe(r.label)}</span><div class="rank-track"><div class="rank-fill" style="width:${Math.max(3,r.count/max*100)}%"></div></div><strong class="rank-value">${fmt.format(r.count)} <small>· ${pct(r.pct)}</small></strong></div>`).join(''):'<p class="muted">Sin datos</p>';
}

function renderExperience(s){
  [['expNps','expNpsText',s.nps],['expPride','expPrideText',s.orgullo],['expSecurity','expSecurityText',s.seguridad],['expInfra','expInfraText',s.infra]].forEach(([id,tid,v])=>{ const card=$(id).closest('.experience-card'); card.classList.remove('red','yellow','green'); card.classList.add(trafficClass(v)); $(id).textContent=one(v); $(tid).textContent=trafficText(v); });
  $('expOffer').textContent=`${fmt.format(s.program_offered_count||0)} · ${pct(s.program_offered_pct||0)}`;
}

function sortStudents(rows){
  const mode=$('studentOrder').value;
  return [...rows].sort((a,b)=>{
    if(mode==='risk') return (a.prob??1)-(b.prob??1);
    if(mode==='prob') return (b.prob??0)-(a.prob??0);
    if(mode==='opp'){ const rank={'Muy alta':5,'Alta':4,'Media-Alta':3,'Media':2,'Baja':1,'Seguimiento':0,'Sin clasificar':0}; return (rank[b.prioridad_convencimiento]||0)-(rank[a.prioridad_convencimiento]||0) || (b.percentil_estructural??0)-(a.percentil_estructural??0) || (a.prob??1)-(b.prob??1); }
    return a.matricula.localeCompare(b.matricula,'es');
  });
}
function visibleStudents(){
  let rows=studentsForCampus();

  const q=$('studentSearch').value.trim().toUpperCase();
  if(q) rows=rows.filter(s=>s.matricula.includes(q));

  if(state.selectedPropensity!=='all'){
    rows=rows.filter(s=>s.propension===state.selectedPropensity);
  }

  const baseFilter=$('baseVivaFilter').value;
  if(baseFilter==='yes') rows=rows.filter(s=>s.en_base_viva);
  else if(baseFilter==='no') rows=rows.filter(s=>!s.en_base_viva);

  const enrolledFilter=$('enrolledFilter').value;
  if(enrolledFilter==='yes') rows=rows.filter(s=>s.inscrito);
  else if(enrolledFilter==='no') rows=rows.filter(s=>!s.inscrito);

  const f=$('opportunityFilter').value;
  if(f==='opportunity'){
    rows=rows.filter(s=>OPPORTUNITY_LEVELS.has(s.prioridad_convencimiento)||['si','sí'].includes(String(s.oportunidad||'').toLowerCase()));
  }else if(['Muy alta','Alta','Media-Alta'].includes(f)){
    rows=rows.filter(s=>s.prioridad_convencimiento===f);
  }else if(f==='other'){
    rows=rows.filter(s=>!OPPORTUNITY_LEVELS.has(s.prioridad_convencimiento)&&!['si','sí'].includes(String(s.oportunidad||'').toLowerCase()));
  }

  return sortStudents(rows);
}

function opportunityStudentsForCampus(){
  // Oportunidades operativas: siguen en Base Viva y todavía no están inscritas.
  return sortStudents(
    studentsForCampus().filter(s=>
      s.es_prospecto &&
      !s.inscrito &&
      (
        OPPORTUNITY_LEVELS.has(s.prioridad_convencimiento) ||
        ['si','sí'].includes(String(s.oportunidad||'').toLowerCase())
      )
    )
  );
}

function csvEscape(v){
  const t=String(v??'');
  return /[",\n]/.test(t)?`"${t.replace(/"/g,'""')}"`:t;
}

function downloadOpportunityCSV(){
  const rows=opportunityStudentsForCampus();
  if(!rows.length){ setExportStatus('No hay matrículas con oportunidad en el filtro de campus actual.'); return; }
  const header=['Matricula','Campus','Estatus','Inscrito','Prioridad_Convencimiento','Propension_Relativa','Potencial_Estructural','Brecha_Oportunidad','Programa_Ofertado'];
  const lines=[header.join(',')];
  rows.forEach(r=>lines.push([
    r.matricula,r.campus,r.estatus_base_viva||'',r.inscrito?'Sí':'No',r.prioridad_convencimiento,r.propension,r.potencial,
    r.brecha_oportunidad==null?'':Number(r.brecha_oportunidad).toFixed(1),
    r.programa_ofertado===1?'Sí':r.programa_ofertado===0?'No':''
  ].map(csvEscape).join(',')));
  const blob=new Blob(['\ufeff'+lines.join('\n')],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a');
  const scope=state.selectedCampus==='__ALL__'?'todos_los_campus':currentSummary().campus_norm.replace(/[^a-z0-9]+/gi,'_');
  a.href=URL.createObjectURL(blob); a.download=`oportunidades_convencimiento_${scope}.csv`; document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},0);
  setExportStatus(`${fmt.format(rows.length)} matrículas descargadas.`);
}

async function copyOpportunityMatriculas(){
  const rows=opportunityStudentsForCampus();
  if(!rows.length){ setExportStatus('No hay matrículas con oportunidad en el filtro de campus actual.'); return; }
  const txt=rows.map(r=>r.matricula).join('\n');
  try{
    if(navigator.clipboard&&window.isSecureContext) await navigator.clipboard.writeText(txt);
    else{
      const ta=document.createElement('textarea');ta.value=txt;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();
    }
    setExportStatus(`${fmt.format(rows.length)} matrículas copiadas.`);
  }catch(e){ setExportStatus('No se pudo copiar automáticamente. Usa Descargar CSV.'); }
}

function setExportStatus(msg){
  const el=$('exportStatus'); if(!el)return; el.textContent=msg||'';
  clearTimeout(setExportStatus._t); setExportStatus._t=setTimeout(()=>{el.textContent='';},5000);
}

function renderStudentList(){
  const rows=visibleStudents(); $('studentCount').textContent=fmt.format(rows.length); const shown=rows.slice(0,300);
  $('studentList').innerHTML=shown.map(s=>`
    <button class="student-row ${state.selectedStudent===s.matricula?'active':''}" data-id="${safe(s.matricula)}">
      <span class="id">${safe(s.matricula)}</span>
      <span class="prob">${pct(s.prob)}</span>
      <span class="campus">${safe(s.campus)}</span>
      <span class="student-status-line">
        <span class="mini-status ${s.inscrito?'enrolled':s.es_prospecto?'yes':'no'}">${safe(s.estatus_base_viva||'Sin match')}</span>
        <span class="mini-status ${s.inscrito?'enrolled':'no'}">${s.inscrito?'Inscrito':'No inscrito'}</span>
      </span>
      <span class="opp-tag ${OPPORTUNITY_LEVELS.has(s.prioridad_convencimiento)?'hot':''}">${safe(s.prioridad_convencimiento||'Seguimiento')}</span>
    </button>
  `).join('')||'<p class="muted">No hay matrículas con los filtros actuales.</p>';
  $('listNote').textContent=rows.length>300?`Mostrando 300 de ${fmt.format(rows.length)}. Usa el buscador para encontrar una matrícula.`:'';
  document.querySelectorAll('.student-row').forEach(el=>el.addEventListener('click',()=>{state.selectedStudent=el.dataset.id;renderStudentList();renderStudentDetail();}));
}
function setText(id,v){$(id).textContent=(v===null||v===undefined||v==='')?'—':v;}
function renderChips(id,items,warning=false){$(id).innerHTML=(items&&items.length)?items.map(x=>`<span class="chip">${safe(x)}</span>`).join(''):'<span class="muted">Sin selección</span>';}
function applyMetricCard(cardId,val){const el=$(cardId);el.classList.remove('red','yellow','green','neutral');el.classList.add(trafficClass(val));}
function renderStudentDetail(){
  const s=state.data.students.find(x=>x.matricula===state.selectedStudent);
  $('studentEmpty').classList.toggle('hidden',!!s); $('studentDetail').classList.toggle('hidden',!s); if(!s)return;
  setText('studentCampus',s.campus);
  setText('studentMatricula',s.matricula);

  const baseBadge=$('studentBaseVivaBadge');
  baseBadge.textContent=s.estatus_base_viva||'Sin match Base Viva';
  baseBadge.className=`badge ${s.inscrito?'enrolled':s.es_prospecto?'live':'inactive'}`;

  const enrolledBadge=$('studentInscritoBadge');
  enrolledBadge.textContent=s.inscrito?'Inscrito':'No inscrito';
  enrolledBadge.className=`badge ${s.inscrito?'enrolled':'inactive'}`;

  setText('studentProb',pct(s.prob)); $('studentProbBar').style.width=`${Math.max(0,Math.min(100,(s.prob||0)*100))}%`; setText('studentProp',`Propensión ${s.propension||'—'}`); setText('studentPct',s.percentil_global==null?'—':`Percentil ${Math.round(s.percentil_global)}`);
  setText('studentPotential',s.potencial); setText('studentPotentialPct',s.percentil_estructural==null?'—':Math.round(s.percentil_estructural)); setText('studentConvPriority',s.prioridad_convencimiento); setText('studentOppGap',s.brecha_oportunidad==null?'—':`${s.brecha_oportunidad>=0?'+':''}${Number(s.brecha_oportunidad).toFixed(1)} pctiles`); setText('studentOpportunity',s.oportunidad);
  setText('studentCareer',s.carrera); setText('studentPlansTec',s.planea_tecmilenio); setText('studentUniversity',s.universidad); setText('studentDesiredCampus',s.campus_interes); setText('studentModality',s.modalidad); setText('studentMobility',s.movilidad); setText('studentStatusBaseViva',s.en_base_viva?(s.estatus_base_viva||'Sin estatus'):'No encontrado en BASE_VIVA.xlsx'); setText('studentPriority',s.prioridad==null?'—':`${s.prioridad} / 5`); setText('studentOffered',s.programa_ofertado===1?'Sí se oferta':s.programa_ofertado===0?'No se oferta':'—'); $('studentOffered').className=s.programa_ofertado===1?'positive':s.programa_ofertado===0?'negative':''; setText('studentCareerCertainty',s.seguridad_carrera==null?'—':`${s.seguridad_carrera} / 5`); setText('studentOrientation',s.orientacion); setText('studentScholarship',scholarship(s.beca)); setText('studentProgramInterest',s.programa_interes);

  setText('studentEnrolledProgram',s.inscrito?s.programa_inscrito:'');
  setText('studentEnrolledCampus',s.inscrito?s.campus_inscrito:'');
  $('studentEnrollmentResult').classList.toggle('not-enrolled',!s.inscrito);
  setText('studentNps',one(s.nps)); setText('studentPride',one(s.orgullo)); setText('studentSecurity',one(s.seguridad_campus)); setText('studentInfra',one(s.infra)); applyMetricCard('studentNpsCard',s.nps); applyMetricCard('studentPrideCard',s.orgullo); applyMetricCard('studentSecurityCard',s.seguridad_campus); applyMetricCard('studentInfraCard',s.infra);
  setText('studentClusterId',s.cluster_id); setText('studentClusterName',s.cluster_nombre); setText('studentClusterDescription',s.cluster_descripcion); setText('studentClusterLever',s.cluster_palanca); renderChips('studentFactors',s.factores); renderChips('studentReasons',s.barreras,true); setText('studentReading',s.lectura_oportunidad); setText('studentSummary',s.resumen_individual);
}

function renderCampusTable(){
  const rows=state.data.campuses; $('campusTableBody').innerHTML=rows.map(s=>{ const high=(s.propensity||[]).filter(x=>['Alta','Muy alta'].includes(x.label)).reduce((a,x)=>a+x.count,0); const highPct=s.survey_n?high/s.survey_n:0; return `<tr data-campus="${safe(s.campus_norm)}"><td><strong>${safe(s.campus)}</strong></td><td class="num">${fmt.format(s.base_viva_n||0)}</td><td class="num">${fmt.format(s.survey_n||0)}</td><td class="num">${s.continuers_expected==null?'—':fmt.format(s.continuers_expected)}</td><td class="num">${fmt.format(s.inscritos_confirmados||0)}</td><td class="num"><strong>${s.forecast_total==null?'—':fmt.format(s.forecast_total)}</strong></td><td class="num">${fmt.format(s.opportunity_count||0)}</td><td><span class="table-state ${cssState(s.state)}">${safe(s.state)}</span></td></tr>`; }).join('');
  document.querySelectorAll('[data-campus]').forEach(tr=>tr.addEventListener('click',()=>{
    state.selectedCampus=tr.dataset.campus;
    state.selectedStudent=null;
    state.selectedPropensity='all';
    $('campusSelect').value=state.selectedCampus;
    $('studentSearch').value='';
    $('opportunityFilter').value='all';
    $('baseVivaFilter').value='all';
    $('enrolledFilter').value='all';
    renderAll();
    window.scrollTo({top:0,behavior:'smooth'});
  }));
}

function renderAll(){
  const s=currentSummary(); renderDataNotes(); renderKPIs(s); renderHistory(s); renderValidation(); renderClusters(s); renderPropensity(); renderRank('topFactors',s.factors); renderRank('topReasons',s.reasons); renderExperience(s); renderStudentList(); renderStudentDetail(); renderCampusTable();
}

function bind(){
  $('tabForecast').addEventListener('click',()=>setDashboardTab('forecast'));
  $('tabValidation').addEventListener('click',()=>setDashboardTab('validation'));

  $('validationThreshold').addEventListener('change',e=>{
    state.validationThreshold=Number(e.target.value)||0.35;
    renderValidation();
  });

  $('campusSelect').addEventListener('change',e=>{
    state.selectedCampus=e.target.value;
    state.selectedStudent=null;
    state.selectedPropensity='all';
    $('studentSearch').value='';
    $('opportunityFilter').value='all';
    $('baseVivaFilter').value='all';
    $('enrolledFilter').value='all';
    renderAll();
  });

  $('studentSearch').addEventListener('input',renderStudentList);
  $('studentOrder').addEventListener('change',renderStudentList);
  $('opportunityFilter').addEventListener('change',renderStudentList);
  $('baseVivaFilter').addEventListener('change',renderStudentList);
  $('enrolledFilter').addEventListener('change',renderStudentList);

  $('filterOpportunitiesBtn').addEventListener('click',()=>{
    state.selectedPropensity='all';
    $('opportunityFilter').value='opportunity';
    $('baseVivaFilter').value='yes';
    $('enrolledFilter').value='no';
    $('studentSearch').value='';
    state.selectedStudent=null;
    renderPropensity();
    renderStudentList();
    renderStudentDetail();
    setExportStatus(`${fmt.format(opportunityStudentsForCampus().length)} oportunidades activas encontradas.`);
  });
  $('copyOpportunitiesBtn').addEventListener('click',copyOpportunityMatriculas);
  $('downloadOpportunitiesBtn').addEventListener('click',downloadOpportunityCSV);
  $('reloadBtn').addEventListener('click',async()=>{try{
    $('loading').classList.remove('hidden');
    state.data=await fetchData(true);
    state.selectedCampus='__ALL__';
    state.selectedStudent=null;
    state.selectedPropensity='all';
    state.selectedDashboardTab='forecast';
    campusOptions();
    $('studentSearch').value='';
    $('opportunityFilter').value='all';
    $('baseVivaFilter').value='all';
    $('enrolledFilter').value='all';
    renderAll();
    setDashboardTab(state.selectedDashboardTab);
  }catch(e){showError(e.message);}finally{$('loading').classList.add('hidden');}});
}
function showError(msg){$('errorBox').textContent=msg;$('errorBox').classList.remove('hidden');setTimeout(()=>$('errorBox').classList.add('hidden'),9000);}

async function init(){
  bind();
  try{
    state.data=await fetchData(false);
    campusOptions();
    renderAll();
  }catch(e){
    showError(e.message);
  }finally{
    $('loading').classList.add('hidden');
  }
}
window.addEventListener('DOMContentLoaded',init);
