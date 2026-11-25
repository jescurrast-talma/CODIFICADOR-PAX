import {normalize} from './excel.js';

const STORAGE_KEY = 'pax-airline-profiles-manual-v2';

function validPairs(pairs) {
  return Array.isArray(pairs) && pairs.every(pair =>
    Array.isArray(pair) && pair.length === 2 &&
    pair.every(value => typeof value === 'string')
  );
}

export function loadProfiles(defaults) {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (!Array.isArray(saved)) return structuredClone(defaults);

    const profiles = saved.filter(profile =>
      profile && typeof profile.id === 'string' &&
      typeof profile.name === 'string' && profile.name.trim() &&
      validPairs(profile.aliases) &&
      (profile.glossary === undefined || validPairs(profile.glossary))
    );

    const merged = new Map(defaults.map(profile => [profile.id, structuredClone(profile)]));
    for (const profile of profiles) {
      const seed = merged.get(profile.id);
      merged.set(profile.id, {...seed, ...profile, group: profile.group ?? seed?.group ?? null});
    }
    return [...merged.values()];
  } catch {
    return structuredClone(defaults);
  }
}

export function saveProfiles(profiles) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
    return true;
  } catch {
    return false;
  }
}

export function createProfile(name, profiles) {
  const cleanName = name.trim().replace(/\s+/g, ' ');
  if (!cleanName || cleanName.length > 60) {
    throw new Error('Escribe un nombre de aerolínea de entre 1 y 60 caracteres.');
  }
  const normalizedName = cleanName.toLocaleLowerCase('es');
  if (profiles.some(profile => profile.name.toLocaleLowerCase('es') === normalizedName)) {
    throw new Error('Esta aerolínea ya tiene un apartado.');
  }
  return {id: crypto.randomUUID(), name: cleanName, aliases: [], group: null};
}

export function profilePairs(profile, basePairs) {
  const overrides = new Set(profile.aliases.map(([value]) => normalize(value)));
  const base = (profile.glossary ?? basePairs).filter(([value]) => !overrides.has(normalize(value)));
  return [...base, ...profile.aliases];
}

export function validateGroup(number, area) {
  const cleanNumber = number.trim();
  const cleanArea = area.trim();
  if (!/^\d{1,15}$/.test(cleanNumber)) throw new Error('El grupo debe ser un número de hasta 15 dígitos.');
  if (!cleanArea || cleanArea.length > 150) throw new Error('Escribe el área del grupo, con un máximo de 150 caracteres.');
  return {number: cleanNumber, area: cleanArea, leader: 'NO'};
}
