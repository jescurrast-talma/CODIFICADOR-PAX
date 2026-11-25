
export function normalize(value) {
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/\b(\d):/g, '0$1:')
    .replace(/(\d{2}:\d{2})\s*\+\s*1\b/g, '$1')
    .replace(/\s*[-–]\s*/g, ' A ');
}

export const specials = [
  ['CURSO CG', 'CAPA'],
  ['SUSPENSION', 'DSUS'],
  ['FALTA', 'F'],
  ['DESCANSO', 'DESD'],
  ['DM', 'DDMS'],
  ['D/M', 'DDMS']
];

const AGENT_HEADER = /^(AGENTES?|APELLIDOS Y NOMBRES|NOMBRES Y APELLIDOS|COLABORADOR(?:ES)?|NOMBRES)$/;
const DAY_HEADER = /^(?:\d{1,2}|\d{4,5}|\d{1,2}[/.-]\d{1,2}(?:[/.-]\d{2,4})?|\d{4}-\d{2}-\d{2}(?:T.*)?)$/;

function decodeXml(text) {
  const entities = {amp: '&', lt: '<', gt: '>', quot: '"', apos: "'"};
  return String(text).replace(/&#(x[\da-f]+|\d+);|&(amp|lt|gt|quot|apos);/gi, (match, number, name) => {
    if (!number) return entities[name.toLowerCase()];
    const codePoint = number[0].toLowerCase() === 'x' ? parseInt(number.slice(1), 16) : Number(number);
    return String.fromCodePoint(codePoint);
  });
}

function attribute(text, name) {
  const match = text.match(new RegExp('(?:^|\\s)' + name + '=["\x27]([^"\x27]*)["\x27]'));
  return match ? decodeXml(match[1]) : '';
}


function xmlElements(xml, name) {
  const expression = new RegExp('<(?:\\w+:)?' + name + '(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</(?:\\w+:)?' + name + '>)', 'g');
  return [...xml.matchAll(expression)].map(match => ({
    attributes: match[1],
    body: match[2] ?? ''
  }));
}

function richText(xml) {
  return xmlElements(xml, 't').map(text => decodeXml(text.body)).join('');
}

function columnNumber(column) {
  return [...column].reduce((number, character) => number * 26 + character.charCodeAt(0) - 64, 0);
}


export async function openExcel(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const readEntry = async path => {
    const entry = zip.file(path);
    return entry ? entry.async('string') : '';
  };
  const workbook = await readEntry('xl/workbook.xml');
  const relationshipsXml = await readEntry('xl/_rels/workbook.xml.rels');
  if (!workbook || !relationshipsXml) throw new Error('Selecciona un archivo Excel .xlsx válido.');

  const relationships = new Map(xmlElements(relationshipsXml, 'Relationship').map(relationship => [
    attribute(relationship.attributes, 'Id'),
    attribute(relationship.attributes, 'Target')
  ]));
  const sharedStrings = xmlElements(await readEntry('xl/sharedStrings.xml'), 'si').map(item => richText(item.body));
  const sheets = [];

  for (const sheet of xmlElements(workbook, 'sheet')) {
    const target = relationships.get(attribute(sheet.attributes, 'r:id'));
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const xml = await readEntry(path);
    if (!/<(?:\w+:)?sheetData\b/.test(xml)) continue;

    const rows = xmlElements(xml, 'row').map(row => xmlElements(row.body, 'c').map(cell => {
      const address = attribute(cell.attributes, 'r');
      const type = attribute(cell.attributes, 't');
      const rawValue = xmlElements(cell.body, 'v')[0]?.body ?? '';
      let value = decodeXml(rawValue);
      if (type === 's') value = sharedStrings[Number(rawValue)] ?? '';
      if (type === 'inlineStr') value = richText(cell.body);
      return {
        address,
        column: address.replace(/\d/g, ''),
        value,
        formula: xmlElements(cell.body, 'f').length > 0
      };
    }));
    sheets.push({name: attribute(sheet.attributes, 'name'), path, xml, rows});
  }
  return {zip, sheets};
}

export function makeLookup(pairs) {
  const lookup = new Map();
  for (const [value, code] of [...pairs, ...specials]) {
    const key = normalize(value);
    const cleanCode = String(code).trim().toUpperCase();
    if (!key || !/^[A-Z0-9_-]+$/.test(cleanCode)) continue;
    if (!lookup.has(key)) lookup.set(key, new Set());
    lookup.get(key).add(cleanCode);
  }
  return lookup;
}

export function readGlossary(book) {
  const pairs = [];
  for (const sheet of book.sheets) {
    const aliasSheet = sheet.rows[0]?.some(cell => cell.column === 'A' && normalize(cell.value) === 'ALIAS');
    for (const row of sheet.rows) {
      let value = row.find(cell => cell.column === 'A')?.value;
      const code = row.find(cell => cell.column === 'B')?.value;
      if (!value || !code) continue;
      const isTime = /\d{1,2}:\d{2}/.test(value);
      const isAlias = aliasSheet && normalize(value) !== 'ALIAS' && /^[A-Z0-9_-]+$/.test(String(code).trim());
      if (!isTime && !isAlias) continue;
      if (value.includes('=')) value = value.split('=').slice(1).join('=').trim();
      pairs.push([value, code]);
    }
  }
  if (!pairs.length) throw new Error('El glosario debe tener horarios o alias en la columna A y códigos en la B.');
  return pairs;
}


export function analyze(book, lookup, overrides = new Map()) {
  const knownCodes = new Set([...lookup.values()].flatMap(set => [...set]));
  const changes = [];
  const pending = new Map();
  const processed = [];

  for (const sheet of book.sheets) {
    let agentColumn = '';
    let dayColumns = [];
    for (const row of sheet.rows) {
      const header = row.find(cell => AGENT_HEADER.test(normalize(cell.value)));
      if (header) {
        agentColumn = header.column;
        dayColumns = row.filter(cell =>
          columnNumber(cell.column) > columnNumber(agentColumn) && DAY_HEADER.test(String(cell.value).trim())
        ).map(cell => cell.column);
        continue;
      }
      if (!agentColumn) continue;
      const agent = row.find(cell => cell.column === agentColumn)?.value;
      if (!agent || /^(TOTAL|LEYENDA|OBSERVACIONES)/.test(normalize(agent))) continue;

      for (const cell of row) {
        if (!dayColumns.includes(cell.column) || cell.formula || !String(cell.value).trim()) continue;
        if (!processed.includes(sheet.name)) processed.push(sheet.name);
        const original = String(cell.value);
        const key = normalize(original);
        if (knownCodes.has(key)) continue;
        const choices = lookup.get(key);
        const code = overrides.get(key) || (choices?.size === 1 ? [...choices][0] : null);
        const item = {sheet: sheet.name, cell, original, agent, code};
        if (code) {
          changes.push(item);
        } else {
          if (!pending.has(key)) pending.set(key, {key, value: original, cells: [], ambiguous: choices?.size > 1});
          pending.get(key).cells.push(item);
        }
      }
    }
  }
  if (!processed.length) {
    throw new Error('No se encontró una tabla de roles. Usa un encabezado AGENTES o APELLIDOS Y NOMBRES y fechas o días a su derecha.');
  }
  return {changes, pending, processed};
}

export async function exportExcel(book, report, type = 'blob') {
  const zip = new JSZip();
  const changesBySheet = new Map();
  for (const change of report.changes) {
    if (!changesBySheet.has(change.sheet)) changesBySheet.set(change.sheet, new Map());
    changesBySheet.get(change.sheet).set(change.cell.address, change.code);
  }
  const modifiedSheets = new Map();

  for (const sheet of book.sheets) {
    const changes = changesBySheet.get(sheet.name);
    if (!changes) continue;
    let count = 0;
    const xml = sheet.xml.replace(/<(?:\w+:)?c(?=[\s/>])([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g, (whole, attributes) => {
      const address = attribute(attributes, 'r');
      const code = changes.get(address);
      if (!code) return whole;
      if (!/^[A-Z0-9_-]+$/.test(code)) throw new Error('Código inválido.');
      count++;
      const cleanAttributes = attributes.replace(/\s+t=["'][^"']*["']/, '');
      return '<c' + cleanAttributes + ' t="inlineStr"><is><t>' + code + '</t></is></c>';
    });
    if (count !== changes.size) throw new Error('No se pudieron localizar todas las celdas para exportar.');
    modifiedSheets.set(sheet.path, xml);
  }

  for (const [path, entry] of Object.entries(book.zip.files)) {
    if (entry.dir) continue;
    zip.file(path, modifiedSheets.has(path) ? modifiedSheets.get(path) : await entry.async('uint8array'));
  }
  return zip.generateAsync({
    type,
    compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}
