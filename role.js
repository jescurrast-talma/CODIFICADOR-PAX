import {exportGroup} from './group.js';
import {cleanName,cleanCode,dateSerial} from './engine.js';
const xml=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function col(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
/** A clean SIGA role: a single header, DNI, AGENTES, and genuine Excel dates. */
export async function exportRole(report,type='blob'){
  if(report.pending.size)throw new Error('Resuelve todos los valores pendientes antes de descargar.');
  const zip=await JSZip.loadAsync(await exportGroup([{name:'TEMP',dni:'00000000'}],{number:'1',area:'PAX'},'uint8array'));
  const rows=[['DNI','AGENTES',...report.days.map(d=>dateSerial(d,report.month,report.year))],...report.records.map(r=>[r.dni,cleanName(r.name),...report.days.map(d=>{
    const c=r.shifts[d];if(!c||c!==cleanCode(c)||!/^(?:[A-Z0-9]+|DLCG\/L)$/.test(c))throw new Error('Código inválido para '+r.name);return c;
  })])];
  const sheetData=rows.map((r,i)=>'<row r="'+(i+1)+'" ht="22" customHeight="1">'+r.map((v,j)=>{
    const style=i===0?(j>=2?2:1):(j>=2?({DESD:3,F:3,DVAC:4,DLCG:5,CAPA:6,DDMS:6}[v]||0):0);
    const a=' r="'+col(j+1)+(i+1)+'" s="'+style+'"';
    return typeof v==='number'?'<c'+a+'><v>'+v+'</v></c>':'<c'+a+' t="inlineStr"><is><t>'+xml(v)+'</t></is></c>';
  }).join('')+'</row>').join('');
  zip.file('xl/worksheets/sheet1.xml','<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:'+col(rows[0].length)+rows.length+'"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="22"/><cols><col min="1" max="1" width="13" customWidth="1"/><col min="2" max="2" width="48" customWidth="1"/><col min="3" max="'+rows[0].length+'" width="12" customWidth="1"/></cols><sheetData>'+sheetData+'</sheetData></worksheet>');
  const colors=['FF002060','FFFF0000','FF00B050','FF00B0F0','FFFFC000'];
  zip.file('xl/styles.xml','<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd-mmm"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'+colors.map(c=>'<fill><patternFill patternType="solid"><fgColor rgb="'+c+'"/><bgColor indexed="64"/></patternFill></fill>').join('')+'</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7">'+[[0,0,0],[0,1,2],[164,1,2],[0,1,3],[0,1,4],[0,0,5],[0,0,6]].map(([n,f,fill])=>'<xf numFmtId="'+n+'" fontId="'+f+'" fillId="'+fill+'" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1" applyFont="1"/>').join('')+'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
  return zip.generateAsync({type,compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
