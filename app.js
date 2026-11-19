const state = { data:null, selectedCampus:'__ALL__', selectedStudent:null, selectedPropensity:'all' };
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
const normalizeId = v => cleanText(v).replace(/\s+/g,'').toUpperCase();
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
    if(!vivaById.has(mid)){
      vivaById.set(mid,{
        campus:displayCampus(campus),
        campus_norm:normalizeCampus(campus)
      });
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
      r['Programa de interes'] ??
      r['Programa de interes'] ??
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

  // Auditoría: en teoría debe ser 0 porque Base Viva ya eliminó
  // a quienes se inscribieron. No modifica la Base Viva.
  let inscritosOverlapViva=0;
  inscritosMatchedById.forEach((_,mid)=>{
    if(vivaById.has(mid)) inscritosOverlapViva++;
  });

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
  // 4) MATCH RESPONDIENTES ∩ BASE VIVA
  // Esta es la ficha y la señal del modelo, igual que antes.
  // =====================================================
  const active=[];
  alumnosAll.forEach(r=>{
    const mid=normalizeId(r['Matrícula']??r['Matricula']);
    if(mid&&vivaById.has(mid)){
      const info=vivaById.get(mid);
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
  vivaById.forEach((info,mid)=>{
    if(!baseIdsByCampus.has(info.campus_norm))baseIdsByCampus.set(info.campus_norm,new Set());
    baseIdsByCampus.get(info.campus_norm).add(mid);
  });

  const inscritosIdsByCampus=new Map();
  inscritosMatchedById.forEach((info,mid)=>{
    if(!inscritosIdsByCampus.has(info.campus_norm))inscritosIdsByCampus.set(info.campus_norm,new Set());
    inscritosIdsByCampus.get(info.campus_norm).add(mid);
  });

  const activeByCampus=groupBy(active,r=>r._campus_norm);

  const campusSet=new Set([
    ...histByCampus.keys(),
    ...baseIdsByCampus.keys(),
    ...activeByCampus.keys(),
    ...inscritosIdsByCampus.keys(),
    ...respondentIdsByCampus.keys()
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

    // Hasta validar nuevamente egresados esperados, la base operativa es Base Viva actual.
    const operationalBase=baseCount||rows.length;
    const baseMethod=baseCount?'Base Viva actual':'respondentes con encuesta/modelo';
    let refRate=null,lowRate=null,highRate=null,rateMethod='sin información suficiente';
    if(hist3!==null&&modelHi!==null){ refRate=(hist3+modelHi)/2; lowRate=Math.min(hist3,modelMean??hist3); highRate=Math.max(hist3,modelHi); rateMethod='punto medio entre promedio 3A ponderado y límite superior del modelo'; }
    else if(modelMean!==null){ refRate=modelMean;lowRate=modelLo??modelMean;highRate=modelHi??modelMean;rateMethod='modelo calibrado (sin histórico comparable)'; }
    else if(hist3!==null){ refRate=lowRate=highRate=hist3;rateMethod='promedio histórico 3A ponderado (sin modelo comparable)'; }
    const cont=refRate===null?null:Math.round(operationalBase*refRate);
    const contLo=lowRate===null?null:Math.round(operationalBase*lowRate);
    const contHi=highRate===null?null:Math.round(operationalBase*highRate);

    // Pronóstico total = pronóstico de la Base Viva actual
    //                  + inscritos confirmados que hacen match con respondientes.
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
    explanation+=` A este pronóstico de Base Viva se suman ${fmt.format(inscritosConfirmados)} inscritos confirmados que hacen match con los respondientes, para un pronóstico total de ${fmt.format(forecastTotal)}.`;
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
      inscrito:!!inscritoInfo,
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


  return {
    meta:{
      survey_total:alumnosAll.length,
      respondent_unique_total:respondentById.size,
      base_viva_total:vivaById.size,
      matched_active:active.length,
      coverage_active:vivaById.size?active.length/vivaById.size:0,
      inscritos_source_unique_total:inscritosSourceIds.size,
      inscritos_respondent_match_total:inscritosMatchedById.size,
      inscritos_overlap_base_viva:inscritosOverlapViva,
      respondents_outside_base_viva:students.filter(s=>!s.en_base_viva).length,
      respondents_not_enrolled:students.filter(s=>!s.inscrito).length,
      analysis_file:'base_continuidad_2027_dashboard.xlsx',
      base_viva_file:'BASE_VIVA.xlsx',
      inscritos_file:'INSCRITOS_2027.xlsx',
      loaded_at:new Date().toISOString(),
      method_note:'El explorador conserva a todos los respondientes de la hoja Alumnos. Cada matrícula se etiqueta según si aparece en BASE_VIVA.xlsx y si aparece en INSCRITOS_2027.xlsx. El pronóstico operativo se mantiene como pronóstico de Base Viva + inscritos con match. Para la gráfica de tendencia, la tasa 2027 usa pronóstico total dividido entre respondientes del dashboard.'
    },
    campuses:campusNorms.map(n=>summaries[n]),
    overall:summaries['__ALL__'],
    students,
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

  $('kpiSurvey').textContent=fmt.format(s.survey_n||0);
  const coverage=s.base_viva_n? s.survey_n/s.base_viva_n : 0;
  $('kpiCoverage').textContent=`${pct(coverage)} de Base Viva con encuesta/modelo`;

  $('kpiBaseForecast').textContent=s.continuers_expected==null?'—':fmt.format(s.continuers_expected);
  $('kpiBaseForecastSub').textContent=`rango ${fmt.format(s.continuers_low||0)}–${fmt.format(s.continuers_high||0)}`;

  $('kpiInscritos').textContent=fmt.format(s.inscritos_confirmados||0);
  $('kpiInscritosSub').textContent='match respondientes ∩ INSCRITOS_2027';

  $('kpiCont').textContent=s.forecast_total==null?'—':fmt.format(s.forecast_total);
  $('kpiContSub').textContent=`${fmt.format(s.continuers_expected||0)} pronóstico Base Viva + ${fmt.format(s.inscritos_confirmados||0)} inscritos`;

  $('kpiOpp').textContent=fmt.format(s.opportunity_count||0);

  const overlap=state.data.meta.inscritos_overlap_base_viva||0;
  const overlapText=overlap?` · ${fmt.format(overlap)} overlap a revisar`:'';
  $('sourceText').textContent=`${fmt.format(state.data.meta.matched_active)} matches Base Viva · ${fmt.format(state.data.meta.inscritos_respondent_match_total)} matches inscritos${overlapText}`;
}

function renderDataNotes(){
  const m=state.data.meta;
  $('surveyNote').innerHTML=`El explorador conserva los <strong>${fmt.format(m.respondent_unique_total)} respondientes únicos</strong> del dashboard. De ellos, <strong>${fmt.format(m.matched_active)}</strong> están hoy en Base Viva y <strong>${fmt.format(m.inscritos_respondent_match_total)}</strong> ya hacen match con INSCRITOS_2027. Los demás también permanecen visibles y quedan etiquetados como “No en Base Viva” y/o “No inscrito”.`;

  $('methodSurvey').textContent=`Pronóstico total = pronóstico actual sobre Base Viva + respondientes que hacen match con INSCRITOS_2027. Para la tendencia 2027, la tasa proyectada se calcula como pronóstico total / respondientes del dashboard. El explorador no elimina registros: conserva todos los respondientes y permite filtrar por propensión y estado operativo.`;
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
      s.en_base_viva &&
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
  const header=['Matricula','Campus','En_Base_Viva','Inscrito','Prioridad_Convencimiento','Propension_Relativa','Potencial_Estructural','Brecha_Oportunidad','Programa_Ofertado'];
  const lines=[header.join(',')];
  rows.forEach(r=>lines.push([
    r.matricula,r.campus,r.en_base_viva?'Sí':'No',r.inscrito?'Sí':'No',r.prioridad_convencimiento,r.propension,r.potencial,
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
        <span class="mini-status ${s.en_base_viva?'yes':'no'}">${s.en_base_viva?'En Base Viva':'No Base Viva'}</span>
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
  baseBadge.textContent=s.en_base_viva?'En Base Viva':'No en Base Viva';
  baseBadge.className=`badge ${s.en_base_viva?'live':'inactive'}`;

  const enrolledBadge=$('studentInscritoBadge');
  enrolledBadge.textContent=s.inscrito?'Inscrito':'No inscrito';
  enrolledBadge.className=`badge ${s.inscrito?'enrolled':'inactive'}`;

  setText('studentProb',pct(s.prob)); $('studentProbBar').style.width=`${Math.max(0,Math.min(100,(s.prob||0)*100))}%`; setText('studentProp',`Propensión ${s.propension||'—'}`); setText('studentPct',s.percentil_global==null?'—':`Percentil ${Math.round(s.percentil_global)}`);
  setText('studentPotential',s.potencial); setText('studentPotentialPct',s.percentil_estructural==null?'—':Math.round(s.percentil_estructural)); setText('studentConvPriority',s.prioridad_convencimiento); setText('studentOppGap',s.brecha_oportunidad==null?'—':`${s.brecha_oportunidad>=0?'+':''}${Number(s.brecha_oportunidad).toFixed(1)} pctiles`); setText('studentOpportunity',s.oportunidad);
  setText('studentCareer',s.carrera); setText('studentPlansTec',s.planea_tecmilenio); setText('studentUniversity',s.universidad); setText('studentDesiredCampus',s.campus_interes); setText('studentModality',s.modalidad); setText('studentMobility',s.movilidad); setText('studentPriority',s.prioridad==null?'—':`${s.prioridad} / 5`); setText('studentOffered',s.programa_ofertado===1?'Sí se oferta':s.programa_ofertado===0?'No se oferta':'—'); $('studentOffered').className=s.programa_ofertado===1?'positive':s.programa_ofertado===0?'negative':''; setText('studentCareerCertainty',s.seguridad_carrera==null?'—':`${s.seguridad_carrera} / 5`); setText('studentOrientation',s.orientacion); setText('studentScholarship',scholarship(s.beca)); setText('studentProgramInterest',s.programa_interes);

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
  const s=currentSummary(); renderDataNotes(); renderKPIs(s); renderHistory(s); renderClusters(s); renderPropensity(); renderRank('topFactors',s.factors); renderRank('topReasons',s.reasons); renderExperience(s); renderStudentList(); renderStudentDetail(); renderCampusTable();
}

function bind(){
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
    campusOptions();
    $('studentSearch').value='';
    $('opportunityFilter').value='all';
    $('baseVivaFilter').value='all';
    $('enrolledFilter').value='all';
    renderAll();
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
