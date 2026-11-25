import {openExcel, readGlossary, makeLookup, analyze, exportExcel, normalize} from './excel.js';
import {loadProfiles, saveProfiles, createProfile, profilePairs, validateGroup} from './profiles.js';
import {exportGroup} from './group.js';
import {createResolver,analyzeRole,readUpdatedGlossary,cleanCode} from './engine.js';
import {exportRole} from './role.js';
let volantes=[];
const period=()=>({month:Number(element('month').value),year:Number(element('year').value)});
const runAnalysis=()=>analyzeRole(book,lookup,period(),overrides);

const element = id => document.getElementById(id);
let basePairs = [];
let profiles = [];
let selectedId = 'ar';
let lookup;
let file = null;
let book = null;
let report = null;
let agents = [];
let overrides = new Map();
let busy = true;

function showMessage(text, error = false) {
  element('message').textContent = text;
  element('message').classList.toggle('error', error);
}

function activeProfile() {
  return profiles.find(profile => profile.id === selectedId);
}

function updateControls() {
  element('convert').disabled = busy || !file || !lookup;
  for (const id of ['role', 'glossary', 'reset', 'apply', 'download', 'add-airline', 'configure-group']) {
    element(id).disabled = busy;
  }
  const groupUnavailable = !activeProfile()?.group || !agents.length || !report || report.pending.size>0;
  element('download').disabled = busy || !report || report.pending.size>0;
  element('month').disabled=busy; element('year').disabled=busy;
  element('download-group').disabled = busy || groupUnavailable;
  element('download-both').disabled = busy || groupUnavailable;
  for (const button of element('airlines').querySelectorAll('button')) button.disabled = busy;
}

function invalidateResult() {
  book = null;
  report = null;
  agents = [];
  overrides.clear();
  element('results').hidden = true;
}

function resetRole() {
  file = null;
  invalidateResult();
  element('role').value = '';
  element('search').value = '';
  element('filename').textContent = 'Arrastra tu Excel aquí';
}

function makeElement(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
}

function renderProfiles() {
  element('airlines').replaceChildren();
  for (const profile of profiles) {
    const button = makeElement('button', profile.name, 'airline-tab');
    button.type = 'button';
    button.setAttribute('aria-current', String(profile.id === selectedId));
    button.addEventListener('click', () => selectProfile(profile.id));
    element('airlines').append(button);
  }
  const profile = activeProfile();
  document.body.dataset.airline=profile.id;
  element("operation-name").textContent=profile.name;
  element("operation-code").textContent=({ar:"AR",aircanada:"AC",airtransat:"TS",arajet:"DM",jetsmart:"JA",sky:"H2",avianca:"AV",delta:"DL",copa:"CM"})[profile.id]||"PAX";
  const pairs = profilePairs(profile, basePairs);
  lookup = createResolver(profile.glossary ?? basePairs, profile, profile.volantes ?? volantes);
  element('glossary-title').textContent = 'Glosario de ' + profile.name;
  element('glossary-label').textContent = pairs.length + ' equivalencias · ' + (profile.glossaryName ?? 'PAX');
  element('airline-description').textContent = 'Configuración de ' + profile.name + '. Las leyendas y el grupo se guardan para este apartado en este navegador.';
  element('group-status').textContent = profile.group ? 'Grupo ' + profile.group.number : 'Configura número y área';
  updateControls();
}

function selectProfile(id) {
  if (busy || id === selectedId) return;
  selectedId = id;
  resetRole();
  element('group-form').hidden = true;
  renderProfiles();
  showMessage('Selecciona el rol de ' + activeProfile().name + '.');
}

function selectRole(selected) {
  if (!selected) return;
  invalidateResult();
  file = null;
  if (!/\.xlsx$/i.test(selected.name) || selected.size > 20 * 1024 * 1024) {
    element('filename').textContent = 'Arrastra tu Excel aquí';
    showMessage('Selecciona un archivo .xlsx de hasta 20 MB.', true);
    updateControls();
    return;
  }
  file = selected;
  element('filename').textContent = selected.name;
  showMessage('Archivo listo para codificar con ' + activeProfile().name + '.');
  updateControls();
}

element('role').addEventListener('change', event => selectRole(event.target.files[0]));
element('drop').addEventListener('dragover', event => {
  event.preventDefault();
  element('drop').classList.add('drag');
});
element('drop').addEventListener('dragleave', () => element('drop').classList.remove('drag'));
element('drop').addEventListener('drop', event => {
  event.preventDefault();
  element('drop').classList.remove('drag');
  if (!busy) selectRole(event.dataTransfer.files[0]);
});

element('add-airline').addEventListener('click', () => {
  element('airline-form').hidden = false;
  element('airline-name').focus();
});
element('cancel-airline').addEventListener('click', () => {
  element('airline-form').hidden = true;
  element('airline-form').reset();
});
element('airline-form').addEventListener('submit', event => {
  event.preventDefault();
  if (busy) return;
  try {
    const profile = createProfile(element('airline-name').value, profiles);
    profiles.push(profile);
    const persisted = saveProfiles(profiles);
    selectedId = profile.id;
    resetRole();
    renderProfiles();
    element('airline-form').hidden = true;
    element('airline-form').reset();
    showMessage(persisted ? 'Apartado creado. Configura su grupo y carga su leyenda cuando corresponda.' : 'Apartado disponible durante esta sesión. El navegador no permite guardar la configuración.');
  } catch (error) {
    showMessage(error.message, true);
  }
});

element('configure-group').addEventListener('click', () => {
  const settings = activeProfile().group;
  element('group-number').value = settings?.number ?? '';
  element('group-area').value = settings?.area ?? '';
  element('group-leader').value = 'NO';
  element('group-form').hidden = false;
  element('group-number').focus();
});
element('cancel-group').addEventListener('click', () => {
  element('group-form').hidden = true;
});
element('group-form').addEventListener('submit', event => {
  event.preventDefault();
  if (busy) return;
  try {
    const settings = validateGroup(element('group-number').value, element('group-area').value);
    settings.leader = 'NO';
    activeProfile().group = settings;
    const persisted = saveProfiles(profiles);
    element('group-form').hidden = true;
    renderProfiles();
    showMessage(persisted ? 'Configuración del grupo guardada.' : 'Grupo configurado durante esta sesión. El navegador no permite guardar la configuración.');
  } catch (error) {
    showMessage(error.message, true);
  }
});

element('glossary').addEventListener('change', async event => {
  const selected = event.target.files[0];
  if (!selected) return;
  busy = true;
  updateControls();
  try {
    if (!/\.xlsx$/i.test(selected.name) || selected.size > 20 * 1024 * 1024) {
      throw new Error('El glosario debe ser un .xlsx de hasta 20 MB.');
    }
    const data = readUpdatedGlossary(await openExcel(await selected.arrayBuffer()));
    const pairs = data.pairs;
    const profile = activeProfile();
    profile.glossary = pairs;
    profile.volantes=data.volantes;
    profile.glossaryName = selected.name;
    const persisted = saveProfiles(profiles);
    invalidateResult();
    renderProfiles();
    showMessage(persisted ? 'Glosario actualizado para ' + profile.name + '.' : 'Glosario aplicado durante esta sesión. El navegador no permite guardar la configuración.');
  } catch (error) {
    showMessage(error.message, true);
  } finally {
    busy = false;
    element('glossary').value = '';
    updateControls();
  }
});

element('convert').addEventListener('click', async () => {
  busy = true;
  updateControls();
  showMessage('Leyendo horarios y preparando los archivos…');
  try {
    book = await openExcel(await file.arrayBuffer());
    report = runAnalysis();
    renderResult();
    try {
      agents = report.records;
      showMessage(report.pending.size ? 'Hay valores pendientes. Resuélvelos para habilitar las descargas.' : activeProfile().group ? 'Rol codificado y grupo listos. ' + agents.length + ' agentes en el grupo.' : 'Rol codificado listo. Configura el grupo de esta aerolínea para generar el segundo archivo.');
    } catch (error) {
      agents = [];
      showMessage('Rol codificado disponible. Para generar el grupo, ' + error.message, true);
    }
    element('results').scrollIntoView({behavior: 'smooth', block: 'start'});
  } catch (error) {
    showMessage(error.message || 'No se pudo leer el Excel.', true);
  } finally {
    busy = false;
    updateControls();
  }
});

function renderResult() {
  element('results').hidden = false;
  element('converted').textContent = report.changes.length.toLocaleString('es-PE');
  element('pending-count').textContent = [...report.pending.values()].reduce((count, entry) => count + entry.cells.length, 0);
  element('sheets-count').textContent = report.processed.length;
  element('pending').hidden = !report.pending.size;
  element('pending-list').replaceChildren();
  for (const entry of report.pending.values()) {
    const row = makeElement('div', undefined, 'pending-row');
    const label = makeElement('label', entry.value + ' — ' + entry.reason + (entry.candidates?.length ? ' ('+entry.candidates.join(', ')+')' : ''));
    const input = makeElement('input');
    input.id = 'pending-' + element('pending-list').children.length;
    input.dataset.key = entry.key;
    input.placeholder = 'Código';
    input.maxLength = 24;
    input.setAttribute('aria-label', 'Código para ' + entry.value);
    label.htmlFor = input.id;
    row.append(label, makeElement('span', entry.cells.length + ' celdas'), input);
    element('pending-list').append(row);
  }
  renderChanges();
}

function renderChanges() {
  if (!report) return;
  const query = normalize(element('search').value);
  const filtered = report.changes.filter(change =>
    normalize([change.original, change.code, change.agent, change.sheet, change.cell.address].join(' ')).includes(query)
  );
  element('changes').replaceChildren();
  for (const change of filtered.slice(0, 300)) {
    const row = makeElement('tr');
    row.append(makeElement('td', change.sheet + ' / ' + change.cell.address), makeElement('td', change.agent), makeElement('td', change.original), makeElement('td', change.code),makeElement('td',change.reason));
    element('changes').append(row);
  }
  element('preview-count').textContent = 'Mostrando ' + Math.min(filtered.length, 300) + ' de ' + filtered.length + ' cambios. El Excel descargado incluye todos.';
}

element('search').addEventListener('input', renderChanges);
element('apply').addEventListener('click', () => {
  const additions = new Map();
  for (const input of element('pending-list').querySelectorAll('input')) {
    const code = cleanCode(input.value);
    if (!code) continue;
    if (!/^(?:[A-Z0-9]{1,24}|DLCG\/L)$/.test(code) || lookup(code).code!==code) {
      showMessage('Selecciona un código existente en el glosario PAX.', true);
      input.focus();
      return;
    }
    additions.set(input.dataset.key, code);
  }
  additions.forEach((value, key) => overrides.set(key, value));
  report = runAnalysis();
  renderResult();
  agents=report.records; updateControls();
  showMessage(report.pending.size ? 'Todavía hay valores pendientes.' : 'Códigos aplicados. Los dos archivos están listos.');
});

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const link = makeElement('a');
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

async function downloadOutput(kind) {
  busy = true;
  updateControls();
  try {
    if(!report || report.pending.size) throw new Error('Resuelve todos los valores pendientes.');
    const stem = activeProfile().id.toUpperCase()+'_'+period().year+'_'+String(period().month).padStart(2,'0');
    if (kind === 'role') {
      downloadBlob(await exportRole(report), stem + '_CODIFICADO.xlsx');
    } else if (kind === 'group') {
      downloadBlob(await exportGroup(agents, activeProfile().group), stem + '_GRUPO.xlsx');
    } else {
      const zip = new JSZip();
      zip.file(stem + '_CODIFICADO.xlsx', await exportRole(report, 'uint8array'));
      zip.file(stem + '_GRUPO.xlsx', await exportGroup(agents, activeProfile().group, 'uint8array'));
      downloadBlob(await zip.generateAsync({type: 'blob', compression: 'DEFLATE'}), stem + '_PARA_CARGA.zip');
    }
    showMessage(report.pending.size ? 'Archivos generados. Revisa los valores pendientes del rol antes de cargarlos al sistema.' : 'Archivos generados.');
  } catch (error) {
    showMessage('No se pudieron generar los archivos. ' + error.message, true);
  } finally {
    busy = false;
    updateControls();
  }
}

element('download').addEventListener('click', () => downloadOutput('role'));
element('download-group').addEventListener('click', () => downloadOutput('group'));
element('download-both').addEventListener('click', () => downloadOutput('both'));
element('reset').addEventListener('click', () => {
  if (busy) return;
  resetRole();
  showMessage('Selecciona tu rol para comenzar.');
  updateControls();
});

try {
  const responses = await Promise.all([fetch('./glossary.json'), fetch('./airlines.json'), fetch('./volantes.json')]);
  if (responses.some(response => !response.ok)) throw new Error('No se pudo cargar la configuración inicial.');
  basePairs = await responses[0].json();
  profiles = loadProfiles(await responses[1].json());
  volantes=await responses[2].json();
  const now=new Date();element('month').value=now.getMonth()+1;element('year').value=now.getFullYear();
  renderProfiles();
} catch (error) {
  showMessage(error.message, true);
} finally {
  busy = false;
  updateControls();
}

for(const id of ['month','year'])element(id).addEventListener('change',()=>{invalidateResult();updateControls();});
