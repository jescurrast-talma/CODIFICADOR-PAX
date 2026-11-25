import {normalize} from './excel.js';
import {cleanName} from './engine.js';

export const GROUP_HEADERS = ['ORDEN DE GRUPO', 'GRUPO', 'NOMBRE', 'DNI', 'LIDER', 'AREA'];


export function collectAgents(book, processedSheets) {
  const agents = new Map();
  for (const sheet of book.sheets) {
    if (!processedSheets.includes(sheet.name)) continue;
    let nameColumn;
    let dniColumn;
    for (const row of sheet.rows) {
      const nameHeader = row.find(cell => /^(AGENTES?|APELLIDOS Y NOMBRES|NOMBRES Y APELLIDOS|COLABORADOR(?:ES)?|NOMBRES)$/.test(normalize(cell.value)));
      const dniHeader = row.find(cell => normalize(cell.value) === 'DNI');
      if (nameHeader && dniHeader) {
        nameColumn = nameHeader.column;
        dniColumn = dniHeader.column;
        continue;
      }
      if (!nameColumn || !dniColumn) continue;
      const name = String(row.find(cell => cell.column === nameColumn)?.value ?? '').trim();
      const dni = String(row.find(cell => cell.column === dniColumn)?.value ?? '').trim();
      if (!name || !dni || /^(TOTAL|LEYENDA|OBSERVACIONES)/.test(normalize(name))) continue;
      if (!/^\d{8}$/.test(dni)) throw new Error('Revisa el DNI de ' + name + '. Debe tener 8 dígitos.');
      const previous = agents.get(dni);
      if (previous && normalize(previous.name) !== normalize(name)) {
        throw new Error('El DNI ' + dni + ' tiene nombres diferentes en el archivo base.');
      }
      if (!previous) agents.set(dni, {name, dni});
    }
  }
  if (!agents.size) throw new Error('No se encontraron agentes con DNI para generar el grupo.');
  return [...agents.values()];
}

export function groupRows(agents, settings) {
  if (!settings || !/^\d{1,15}$/.test(settings.number) || !settings.area) {
    throw new Error('Configura el número y el área del grupo de esta aerolínea.');
  }
  return [GROUP_HEADERS, ...agents.map((agent, index) => [
    index + 1, Number(settings.number), cleanName(agent.name),
    agent.dni.startsWith('0') ? agent.dni : Number(agent.dni),
    'NO', cleanName(settings.area)
  ])];
}

function escapeXml(value) {
  return String(value).replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;'
  })[character]);
}


export async function exportGroup(agents, settings, type = 'blob') {
  const rows = groupRows(agents, settings);
  const namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const sheetRows = rows.map((row, rowIndex) => {
    const cells = row.map((value, columnIndex) => {
      const address = String.fromCharCode(65 + columnIndex) + (rowIndex + 1);
      const style = rowIndex === 0 ? ' s="1"' : '';
      if (typeof value === 'number') return '<c r="' + address + '"' + style + '><v>' + value + '</v></c>';
      return '<c r="' + address + '" t="inlineStr"' + style + '><is><t xml:space="preserve">' + escapeXml(value) + '</t></is></c>';
    }).join('');
    return '<row r="' + (rowIndex + 1) + '" ht="22" customHeight="1">' + cells + '</row>';
  }).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="' + namespace + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Hoja1" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="' + namespace + '"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF136C5E"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="' + namespace + '"><dimension ref="A1:F' + rows.length + '"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="22"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="3" width="55" customWidth="1"/><col min="4" max="5" width="14" customWidth="1"/><col min="6" max="6" width="42" customWidth="1"/></cols><sheetData>' + sheetRows + '</sheetData></worksheet>');
  return zip.generateAsync({type, compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
