import {normalize, makeLookup, specials} from './excel.js';

export const cleanName = value => String(value ?? '').replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\s+/gu, ' ').trim();
export const cleanCode = value => cleanName(value).replace(/\s/g, '').toUpperCase();
const aliasKey = value => cleanCode(value).replace(/[.!,;:*]+$/g, '');
const nameHeader = /^(AGENTES?|APELLIDOS Y NOMBRES|NOMBRES Y APELLIDOS|COLABORADOR(?:ES)?|NOMBRES|NOMBRE)$/;
const timeKey = value => normalize(String(value).replace(/\(\s*\+\s*1\s*\)/g, '').replace(/—/g, '-'));
const isTime = value => /^\d{2}:\d{2} A \d{2}:\d{2}$/.test(timeKey(value));
const single = set => set?.size === 1 ? [...set][0] : null;
function minutes(time) { const [h,m]=time.split(':').map(Number); return h*60+m; }
function timeVector(value) { return [...value.matchAll(/\d{2}:\d{2}/g)].map(m => minutes(m[0])); }

export function createResolver(pairs, profile, volantes = []) {
  const cleanPairs=pairs.map(([v,c])=>[v,cleanCode(c)]);
  const lookup=makeLookup(cleanPairs);
  const known=new Set([...cleanPairs.map(p=>p[1]),...specials.map(p=>p[1]),...volantes.map(v=>v.code)]);
  
  for(const c of ['DVAC','DLCG']) known.add(c);
  const aliases=new Map((profile.aliases||[]).map(([a,b])=>[aliasKey(a),b]));
  const byCode=new Map(cleanPairs.filter(([v])=>isTime(v)).map(([v,c])=>[c,timeKey(v)]));
  const exact=value=>single(lookup.get(timeKey(value)));
  return function resolve(value) {
    const codeInput=cleanCode(value), key=timeKey(value), ak=aliasKey(value);
    if(known.has(codeInput)) return {code:codeInput,reason:'Código PAX'};
    if(profile.id==='jetsmart' && /^ACT\d*$/.test(ak)) return {code:'DESD',reason:'ACT Jet Smart'};
    if(profile.id==='copa' && isTime(value)) return {code:'DESD',reason:'Horario numérico Copa'};
    let alias=ak;
    if(profile.id==='copa' && /^OJT(?:-|\s|$)/i.test(cleanName(value))) {
      const suffix=cleanName(value).replace(/^OJT[\s-]*/i,'');
      if(!suffix) return {code:'CAPA',reason:'OJT sin turno'};
      alias=aliasKey(suffix);
    }
    if(profile.unresolved?.[alias]) return {reason:'Leyenda contradictoria',candidates:profile.unresolved[alias].map(exact).filter(Boolean)};
    if(aliases.has(alias)) {
      const target=aliases.get(alias),code=known.has(cleanCode(target))?cleanCode(target):exact(target);
      return code?{code,reason:'Leyenda '+profile.name}:{reason:'Horario de la leyenda sin equivalencia PAX'};
    }
    
    if(/^[A-Z]+\d+[A-Z]*\/[A-Z]+\d+[A-Z]*$/.test(codeInput)) {
      const explicit=volantes.filter(v=>(v.aliases||[]).map(cleanCode).includes(codeInput));
      if(explicit.length===1)return {code:explicit[0].code,reason:'Equivalencia VOLANTE'};
      const components=codeInput.split('/').map(c=>byCode.get(c));
      if(components.some(v=>!v))return {reason:'Código componente no encontrado'};
      const target=components.join(' Y '), vec=timeVector(target);
      const scored=volantes.map(v=>({v,dist:timeVector(v.time).reduce((sum,t,i)=>sum+Math.min(Math.abs(t-vec[i]),1440-Math.abs(t-vec[i])),0)})).sort((a,b)=>a.dist-b.dist);
      if(!scored.length)return {reason:'Glosario sin volantes'};
      const closest=scored.filter(s=>s.dist===scored[0].dist);
      if(closest.length!==1)return {reason:'Empate entre volantes cercanos',candidates:closest.map(s=>s.v.code)};
      return {code:closest[0].v.code,reason:closest[0].dist?'Volante más cercano: '+target+' → '+closest[0].v.time:'Volante exacto',approximate:closest[0].dist>0};
    }
    if(/^CURSO\b/.test(key))return {code:'CAPA',reason:'Curso'};
    const code=exact(value);
    if(code)return {code,reason:'Glosario PAX'};
    return {reason:lookup.get(key)?.size>1?'Equivalencia ambigua':'Sin equivalencia en glosario',candidates:[...(lookup.get(key)||[])]};
  };
}

export function readUpdatedGlossary(book) {
  const pairs=[],volantes=[];
  for(const sheet of book.sheets) {
    for(const row of sheet.rows) {
      const at=c=>row.find(x=>x.column===c)?.value;
      const a=at('A'),b=at('B');
      if(!a||!b||cleanCode(b)==='CODIGO')continue;
      if(/VOLANTE/i.test(sheet.name)) {
        if(!/^[A-Z0-9]+$/.test(cleanCode(b)))continue;
        volantes.push({time:String(a).split('=').pop().trim(),code:cleanCode(b),aliases:[at('C'),at('D')].filter(Boolean)});
      } else if(/^[A-Z0-9]+$/.test(cleanCode(b)))pairs.push([cleanName(a),cleanCode(b)]);
    }
  }
  if(!pairs.length)throw new Error('No se encontraron equivalencias en las columnas HORARIO y CODIGO.');
  return {pairs,volantes};
}

const excelEpoch=Date.UTC(1899,11,30);
export function dateSerial(day,month,year){return Math.round((Date.UTC(year,month-1,day)-excelEpoch)/86400000);}
function parseDay(value,month,year) {
  const s=cleanName(value);let day,m=month,y=year;
  if(/^\d{1,2}$/.test(s))day=Number(s);
  else if(/^\d{5}(?:\.0+)?$/.test(s)) {const d=new Date(excelEpoch+Number(s)*86400000);day=d.getUTCDate();m=d.getUTCMonth()+1;y=d.getUTCFullYear();}
  else if(/^\d{4}-\d{2}-\d{2}/.test(s)){[y,m,day]=s.slice(0,10).split('-').map(Number);}
  else {
    const match=s.match(/^(\d{1,2})[/.\-](\d{1,2})(?:[/.\-](\d{2,4}))?$/);
    if(match){day=+match[1];m=+match[2];if(match[3])y=+match[3]+(match[3].length===2?2000:0);}
    else {const match=s.toLowerCase().match(/^(\d{1,2})[- /](ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)(?:[- /](\d{4}))?$/);if(!match)return null;day=+match[1];m=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'].indexOf(match[2])+1;if(match[3])y=+match[3];}
  }
  if(day<1||day>new Date(Date.UTC(y,m,0)).getUTCDate()||m<1||m>12)throw new Error('Fecha inválida: '+s);
  if(m!==month||y!==year)throw new Error('La fecha '+s+' no corresponde al mes y año seleccionados.');
  return day;
}

export function analyzeRole(book, resolver, {month,year}, overrides=new Map()) {
  month=Number(month);year=Number(year);
  if(!Number.isInteger(month)||month<1||month>12||!Number.isInteger(year)||year<2000||year>2100)throw new Error('Selecciona un mes y año válidos.');
  const records=[],byDni=new Map(),dates=new Set(),processed=[],changes=[],pending=new Map(),warnings=[];
  for(const sheet of book.sheets) {
    let nc,dc,days=[];
    for(const row of sheet.rows) {
      const nh=row.find(c=>nameHeader.test(normalize(c.value))), dh=row.find(c=>normalize(c.value)==='DNI');
      if(nh&&dh){
        nc=nh.column;dc=dh.column;days=[];
        for(const c of row){if(c===nh||c===dh)continue;const day=parseDay(c.value,month,year);if(day!==null){if(days.some(d=>d.day===day))throw new Error('Día duplicado en '+sheet.name);days.push({column:c.column,day});dates.add(day);}}
        continue;
      }
      if(!nc||!days.length)continue;
      const name=cleanName(row.find(c=>c.column===nc)?.value), dni=cleanCode(row.find(c=>c.column===dc)?.value);
      if(!name&&!dni)continue;
      if(/^(TOTAL|LEYENDA|OBSERVACIONES)/.test(normalize(name))){nc=null;continue;}
      // Ignore decorations with neither a DNI nor shift data.
      if(!dni&&!row.some(c=>days.some(d=>d.column===c.column)&&cleanName(c.value)))continue;
      if(!name||!/^\d{8}$/.test(dni))throw new Error('Revisa nombre y DNI (8 dígitos) de '+(name||dni)+' en '+sheet.name+'.');
      if(!processed.includes(sheet.name))processed.push(sheet.name);
      let record=byDni.get(dni);
      if(record && normalize(record.name)!==normalize(name))throw new Error('El DNI '+dni+' tiene nombres diferentes.');
      if(!record){record={name,dni,shifts:{}};records.push(record);byDni.set(dni,record);}
      for(const d of days){
        const cell=row.find(c=>c.column===d.column)||{column:d.column,address:d.column+(row[0]?.address?.match(/\d+/)?.[0]||'?'),value:''};
        const original=String(cell.value??''),key=timeKey(original)||'(VACÍO)',resolution=resolver(original);
        const override=overrides.get(key);
        if(override){const checked=resolver(override);if(!checked.code||checked.code!==cleanCode(override))throw new Error('El código manual '+override+' no está en el glosario.');resolution.code=checked.code;resolution.reason='Corrección manual';}
        if(cell.formula){resolution.code=null;resolution.reason='Fórmula: reemplázala por un valor en el archivo base';}
        const item={sheet:sheet.name,cell,original,agent:name,code:resolution.code,reason:resolution.reason,day:d.day};
        if(Object.hasOwn(record.shifts,d.day))throw new Error('DNI '+dni+' repetido para el día '+d.day+'. Revisa las filas duplicadas.');
        record.shifts[d.day]=resolution.code||null;
        if(!resolution.code){
          if(!pending.has(key))pending.set(key,{key,value:original||'(Vacío)',cells:[],reason:resolution.reason,candidates:resolution.candidates||[]});
          pending.get(key).cells.push(item);
        }else if(original!==resolution.code)changes.push(item);
        if(resolution.approximate)warnings.push(item);
      }
    }
  }
  if(!records.length)throw new Error('No se encontró una tabla con DNI, AGENTES y fechas o días en la misma fila.');
  const days=[...dates].sort((a,b)=>a-b);
  for(const r of records)for(const d of days)if(!Object.hasOwn(r.shifts,d))throw new Error('Falta el día '+d+' para '+r.name+'.');
  return {records,days,month,year,changes,pending,processed,warnings};
}
