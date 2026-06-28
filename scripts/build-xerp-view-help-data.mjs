import fs from 'node:fs';
import path from 'node:path';

const workspaceRoot = 'C:/Users/micha/Documents/X-ERP-HELP';
const xerpRoot = 'C:/X-ERP/X-ERP';
const viewsRoot = path.join(xerpRoot, 'X-ERP.Client/Pages/Views');
const sharedRoot = path.join(xerpRoot, 'X-ERP.Shared');
const resxPath = path.join(xerpRoot, 'X-ERP.Client/ResourceFiles/LzrResource.de-DE.resx');
const registerPath = path.join(sharedRoot, 'Constants/SystemRegisterData.cs');
const outputDir = path.join(workspaceRoot, 'outputs/help-seo');

const fieldTags = new Set([
  'XObId', 'XText', 'XSele', 'XQuan', 'XChec', 'XInte', 'XLook', 'XDate', 'XMemo',
  'XDrop', 'XTime', 'XCalc', 'XSpin', 'XColo', 'XHtml', 'XFile', 'XYear', 'XPeri',
  'XRadi', 'XPass', 'XEMail', 'XUrl', 'XDeci', 'XBool', 'DxFormLayoutItem',
]);

function walk(dir, predicate = () => true, acc = []) {
  if (!fs.existsSync(dir)) return acc;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, predicate, acc);
    else if (predicate(full)) acc.push(full);
  }
  return acc;
}

function decodeXml(s) {
  return String(s ?? '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .trim();
}

function read(file) {
  return fs.readFileSync(file, 'utf8');
}

function loadTranslations() {
  const text = read(resxPath);
  const map = new Map();
  const rx = /<data\s+name="([^"]+)"[\s\S]*?<value>([\s\S]*?)<\/value>/g;
  for (const m of text.matchAll(rx)) {
    map.set(m[1], decodeXml(m[2]));
  }
  return map;
}

function loadRegisters(translations) {
  const text = read(registerPath);
  const byType = new Map();
  const byId = new Map();
  const rx = /new\("([^"]+)",\s*(\d+),\s*"([^"]+)",\s*"([^"]+)"\)/g;
  for (const m of text.matchAll(rx)) {
    const entry = {
      id: m[1],
      order: Number(m[2]),
      name: m[3],
      deName: translate(m[3], translations),
      type: m[4],
    };
    if (!byType.has(entry.type.toLowerCase())) byType.set(entry.type.toLowerCase(), []);
    byType.get(entry.type.toLowerCase()).push(entry);
    byId.set(entry.id, entry);
  }
  for (const list of byType.values()) list.sort((a, b) => a.order - b.order);
  return { byType, byId };
}

function translate(text, translations) {
  if (!text) return '';
  const clean = String(text).trim();
  return translations.get(clean) || clean;
}

function stripComputedPrefix(label) {
  return String(label ?? '').replace(/^[\s:=+\-*]+/, '').trim();
}

function splitWords(s) {
  return String(s ?? '')
    .replace(/([a-zäöüß])([A-ZÄÖÜ])/g, '$1 $2')
    .replace(/[_-]+/g, ' ')
    .trim();
}

function slugify(s) {
  const replacements = new Map([
    ['ä', 'ae'], ['ö', 'oe'], ['ü', 'ue'], ['ß', 'ss'],
    ['Ä', 'ae'], ['Ö', 'oe'], ['Ü', 'ue'],
  ]);
  return String(s ?? '')
    .replace(/[äöüßÄÖÜ]/g, ch => replacements.get(ch) || ch)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-{2,}/g, '-') || 'seite';
}

function csvEscape(value) {
  const s = String(value ?? '');
  if (/[;"\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function writeTable(name, rows) {
  fs.mkdirSync(outputDir, { recursive: true });
  const jsonPath = path.join(outputDir, `${name}.json`);
  const csvPath = path.join(outputDir, `${name}.csv`);
  fs.writeFileSync(jsonPath, JSON.stringify(rows, null, 2), 'utf8');
  const headers = Array.from(rows.reduce((set, row) => {
    for (const key of Object.keys(row)) set.add(key);
    return set;
  }, new Set()));
  const csv = [
    headers.map(csvEscape).join(';'),
    ...rows.map(row => headers.map(h => csvEscape(row[h])).join(';')),
  ].join('\r\n');
  fs.writeFileSync(csvPath, csv, 'utf8');
}

function attrValue(tag, attr) {
  const lzrRx = new RegExp(`${attr}\\s*=\\s*"?@?\\(?\\s*lzr\\[\\s*"([^"]+)"\\s*\\]\\s*\\)?"?`, 'i');
  const lzr = lzrRx.exec(tag);
  if (lzr) return `@lzr["${lzr[1]}"]`;
  const rx = new RegExp(`${attr}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|@([^\\s>]+))`, 'i');
  const m = rx.exec(tag);
  return m ? (m[1] ?? m[2] ?? m[3] ?? '') : '';
}

function lzrKey(s) {
  const m = /lzr\[\s*"([^"]+)"\s*\]/.exec(s);
  return m?.[1] || '';
}

function labelFromAttr(raw, translations) {
  if (!raw) return '';
  const key = lzrKey(raw);
  return translate(key || raw.replace(/^@/, '').trim(), translations);
}

function bindField(tag) {
  const raw = attrValue(tag, 'BindField');
  if (!raw) return '';
  const nameof = /nameof\((?:[A-Za-z_][A-Za-z0-9_]*\.)?([A-Za-z_][A-Za-z0-9_]*)\)/.exec(raw);
  if (nameof) return nameof[1];
  return raw.replace(/^@/, '').replace(/^nameof\(/, '').replace(/\)$/, '').split('.').pop();
}

function findRegisterConstants(text) {
  const constants = new Map();
  const direct = /const\s+string\s+(REGISTER_\w+)\s*=\s*"([^"]+)"/g;
  for (const m of text.matchAll(direct)) constants.set(m[1], m[2]);
  const aliases = new Map();
  const aliasRx = /const\s+string\s+(REGISTER_\w+)\s*=\s*(REGISTER_\w+)\s*;/g;
  for (const m of text.matchAll(aliasRx)) aliases.set(m[1], m[2]);
  for (let i = 0; i < 10; i++) {
    let changed = false;
    for (const [key, target] of aliases) {
      if (!constants.has(key) && constants.has(target)) {
        constants.set(key, constants.get(target));
        changed = true;
      } else if (aliases.has(target)) {
        aliases.set(key, aliases.get(target));
        changed = true;
      }
    }
    if (!changed) break;
  }
  return constants;
}

function registerTypeFrom(text, entityName, viewName) {
  const direct = /SystemRegisterData\.GetByRegisterType\("([^"]+)"\)/.exec(text);
  if (direct) return direct[1];
  if (viewName.endsWith('Wizard')) return viewName;
  return entityName || viewName.replace(/Edit$/, '').replace(/Wizard$/, '');
}

function findKind(file, text) {
  const base = path.basename(file, '.razor');
  if (/Wizard$/.test(base) || /<XTBWizard\b/.test(text)) return 'Wizard';
  if (/@inherits\s+XViewEdit\b/.test(text) || /Edit$/.test(base)) return 'EditView';
  if (/List$/.test(base) || /ListMobile$/.test(base) || /ListSelect$/.test(base)) return 'ListView';
  return 'View';
}

function entityNameFrom(text, viewName) {
  const m = /@inherits\s+XViewEdit\s*<\s*([^,\s>]+)/.exec(text);
  if (m) return m[1].trim();
  return viewName.replace(/Edit$/, '').replace(/Wizard$/, '').replace(/List$/, '');
}

function componentMap(razorFiles) {
  const map = new Map();
  for (const file of razorFiles) map.set(path.basename(file, '.razor'), file);
  return map;
}

function extractEvents(file, text, translations, components, depth = 0, offset = 0, seen = new Set()) {
  const events = [];
  const constants = findRegisterConstants(text);

  for (const m of text.matchAll(/SelectedRegisterId\s*==\s*(REGISTER_\w+)/g)) {
    events.push({ pos: offset + m.index, type: 'REGISTER', registerId: constants.get(m[1]) || m[1] });
  }
  for (const m of text.matchAll(/SelectedRegisterId\s*==\s*"([^"]+)"/g)) {
    events.push({ pos: offset + m.index, type: 'REGISTER', registerId: m[1] });
  }
  for (const m of text.matchAll(/register\.Id\s*==\s*"([^"]+)"/g)) {
    events.push({ pos: offset + m.index, type: 'REGISTER', registerId: m[1] });
  }

  for (const m of text.matchAll(/<([A-Za-z][A-Za-z0-9_]*)\b[\s\S]*?(?:\/>|>)/g)) {
    const tag = m[0];
    const tagName = m[1];
    const pos = offset + m.index;
    if (fieldTags.has(tagName) && /\bCaption\s*=/.test(tag)) {
      const label = labelFromAttr(attrValue(tag, 'Caption'), translations);
      if (label) {
        const computed = /^[\s:=+\-*]/.test(label);
        events.push({
          pos,
          type: computed ? 'DISPLAY_FIELD' : 'FIELD',
          label: computed ? stripComputedPrefix(label) : label,
          rawLabel: label,
          bindField: bindField(tag),
          control: tagName,
          disabled: /\bDisabled\s*=\s*(?:true|"true"|@true)/i.test(tag) ? 'true' : '',
        });
      }
    }
    if (/List(?:Mobile|Select)?$/.test(tagName) && !/^X/.test(tagName)) {
      events.push({
        pos,
        type: 'LIST',
        label: splitWords(tagName.replace(/List(?:Mobile|Select)?$/, ' Liste')),
        control: tagName,
      });
    }
    if (tagName === 'DxButton') {
      const label = labelFromAttr(attrValue(tag, 'Text'), translations);
      events.push({ pos, type: 'BUTTON', label: label || 'Schaltflaeche', control: tagName });
    }
    if (tagName === 'XTBEdit') {
      const standard = ['Neu', 'Speichern', 'Speichern und schlie\u00dfen', 'Abbrechen oder schlie\u00dfen'];
      if (/Preview_Clicked\s*=/.test(tag)) standard.push('Vorschau');
      if (/OnFieldChooserClicked\s*=/.test(tag)) standard.push('Feldauswahl');
      if (/OnFieldResetClicked\s*=/.test(tag)) standard.push('Feldlayout zur\u00fccksetzen');
      standard.forEach((label, idx) => events.push({ pos: pos + idx / 100, type: 'BUTTON', label, control: tagName }));
    }
    if (tagName === 'XTBWizard') {
      ['Zur\u00fcck', 'Weiter', 'Fertigstellen oder ausf\u00fchren', 'Abbrechen'].forEach((label, idx) => {
        events.push({ pos: pos + idx / 100, type: 'BUTTON', label, control: tagName });
      });
    }
    if (depth < 2 && components.has(tagName) && components.get(tagName) !== file) {
      const child = components.get(tagName);
      const key = `${file}->${child}`;
      if (!seen.has(key)) {
        seen.add(key);
        events.push(...extractEvents(child, read(child), translations, components, depth + 1, pos + 0.01, seen));
      }
    }
  }

  for (const m of text.matchAll(/<a\b[^>]*\bbutton\b[^>]*>([\s\S]*?)<\/a>/g)) {
    const inner = m[1];
    const key = lzrKey(inner);
    const cleaned = inner.replace(/<[^>]+>/g, '').replace(/@\(?lzr\["[^"]+"\]\)?/g, '').trim();
    const label = translate(key || cleaned || 'Schaltfl\u00e4che', translations);
    events.push({ pos: offset + m.index, type: 'BUTTON', label, control: 'a[button]' });
  }

  return events.sort((a, b) => a.pos - b.pos);
}

function uniquePush(list, seen, row, key) {
  if (seen.has(key)) return;
  seen.add(key);
  list.push(row);
}

function viewGermanName(viewName, kind, entityName, translations) {
  const entity = translate(entityName, translations);
  if (kind === 'EditView') return `${entity} bearbeiten`;
  if (kind === 'Wizard') return `${entity} Assistent`;
  if (kind === 'ListView') return `${entity} Liste`;
  return translate(splitWords(viewName), translations);
}

function textDraft(viewName, kind, deName, moduleName, tabs, fields, buttons) {
  const tabText = tabs.length
    ? `Die Ansicht ist in ${tabs.length} Registertabs gegliedert: ${tabs.slice(0, 8).map(t => t.deName).join(', ')}${tabs.length > 8 ? ' und weitere' : ''}.`
    : 'Die Ansicht ist ohne Registertabs aufgebaut.';
  const purpose = kind === 'Wizard'
    ? `${deName} ist ein ERP-Assistent in X-ERP. Er f\u00fchrt Anwender schrittweise durch einen abgegrenzten Prozess im Bereich ${moduleName}.`
    : `${deName} ist eine ERP-Bearbeitungsansicht in X-ERP. Sie dient dazu, Datens\u00e4tze im Bereich ${moduleName} anzulegen, zu pr\u00fcfen und zu bearbeiten.`;
  return `${purpose} ${tabText} Die Hilfeseite sollte zuerst den Zweck der Ansicht erkl\u00e4ren, danach die wichtigsten Felder beschreiben und anschlie\u00dfend die verf\u00fcgbaren Schaltfl\u00e4chen nennen. In der aktuellen Quellstruktur wurden ${fields} Felder/Anzeigewerte und ${buttons} Schaltfl\u00e4chen erkannt.`;
}

const translations = loadTranslations();
const registers = loadRegisters(translations);
const allRazorFiles = walk(xerpRoot, f => f.endsWith('.razor')).sort();
const razorFiles = walk(viewsRoot, f => f.endsWith('.razor')).sort();
const csFiles = walk(xerpRoot, f => f.endsWith('.cs')).sort();
const components = componentMap(razorFiles);

let razorBytesRead = 0;
for (const file of allRazorFiles) {
  razorBytesRead += Buffer.byteLength(read(file), 'utf8');
}

let csBytesRead = 0;
let csApiClientHits = 0;
for (const file of csFiles) {
  const text = read(file);
  csBytesRead += Buffer.byteLength(text, 'utf8');
  if (/ApiClient|XView|SystemRegisterData|Wizard|Edit/.test(text)) csApiClientHits++;
}

const inventory = [];
const structure = [];
const drafts = [];

for (const file of razorFiles) {
  const text = read(file);
  if (!/@page\s+"/.test(text) && !/@inherits\s+XViewEdit\b/.test(text) && !/<XTBWizard\b/.test(text)) continue;

  const viewName = path.basename(file, '.razor');
  const rel = path.relative(viewsRoot, file).replace(/\\/g, '/');
  const moduleName = rel.split('/')[0];
  const kind = findKind(file, text);
  const entityName = entityNameFrom(text, viewName);
  const registerType = registerTypeFrom(text, entityName, viewName);
  const deName = viewGermanName(viewName, kind, entityName, translations);
  const routes = [...text.matchAll(/@page\s+"([^"]+)"/g)].map(m => m[1]).join(' | ');
  const registerList = registers.byType.get(registerType.toLowerCase()) || [];
  const events = (kind === 'EditView' || kind === 'Wizard')
    ? extractEvents(file, text, translations, components)
    : [];

  const counts = { FIELD: 0, DISPLAY_FIELD: 0, BUTTON: 0, LIST: 0 };
  const byRegister = new Map();
  const rootKey = `${viewName}::__ROOT__`;
  byRegister.set('__ROOT__', { id: '', name: viewName, deName, events: [] });
  for (const reg of registerList) byRegister.set(reg.id, { ...reg, events: [] });

  let currentRegister = '__ROOT__';
  for (const event of events) {
    if (event.type === 'REGISTER') {
      currentRegister = byRegister.has(event.registerId) ? event.registerId : event.registerId;
      if (!byRegister.has(currentRegister)) {
        const inferred = registers.byId.get(currentRegister);
        byRegister.set(currentRegister, inferred ? { ...inferred, events: [] } : {
          id: currentRegister,
          order: 999999,
          name: currentRegister.split('-').slice(1).join('-') || currentRegister,
          deName: translate(currentRegister.split('-').slice(1).join(' '), translations),
          type: registerType,
          events: [],
        });
      }
      continue;
    }
    if (counts[event.type] !== undefined) counts[event.type]++;
    byRegister.get(currentRegister)?.events.push(event);
  }

  inventory.push({
    VIEW_NAME: viewName,
    DE_NAME: deName,
    KIND: kind,
    MODULE: moduleName,
    ENTITY: entityName,
    REGISTER_TYPE: registerType,
    ROUTES: routes,
    RAZOR_FILE: file,
    REGISTER_COUNT: registerList.length,
    FIELD_COUNT: counts.FIELD,
    DISPLAY_FIELD_COUNT: counts.DISPLAY_FIELD,
    LIST_COUNT: counts.LIST,
    BUTTON_COUNT: counts.BUTTON,
    TITLE: `ERP Hilfe: ${deName} in X-ERP`,
    META_DESCRIPTION: `${deName}: Zweck, Aufbau, Felder und Schaltflaechen der ERP-Ansicht in X-ERP verstaendlich erklaert.`,
    DIRECTORY_PATH: `views/${slugify(moduleName)}/${slugify(viewName)}`,
    URL_PATH: `/de/help/views/${slugify(moduleName)}/${slugify(viewName)}/`,
  });

  if (kind !== 'EditView' && kind !== 'Wizard') continue;

  const rootSlug = `views/${slugify(moduleName)}/${slugify(viewName)}`;
  uniquePush(structure, new Set(), {}, 'unused');
  structure.pop();
  const seen = new Set();
  const rootPageId = `views/${viewName}`;
  uniquePush(structure, seen, {
    PAGE_ID: rootPageId,
    PARENT_PAGE_ID: '',
    LEVEL: 0,
    SORT: 0,
    TYPE: kind,
    VIEW_NAME: viewName,
    REGISTER_ID: '',
    ITEM_KEY: viewName,
    DE_NAME: deName,
    SOURCE_NAME: viewName,
    BIND_FIELD: '',
    CONTROL: '',
    DIRECTORY_PATH: rootSlug,
    URL_PATH: `/de/help/views/${slugify(moduleName)}/${slugify(viewName)}/`,
    TITLE: `ERP Hilfe: ${deName} in X-ERP`,
    META_DESCRIPTION: `${deName}: Zweck, Aufbau, Felder und Schaltflaechen der ERP-Ansicht in X-ERP.`,
  }, rootPageId);

  const orderedRegisters = [
    ...registerList.map(r => byRegister.get(r.id)).filter(Boolean),
    ...[...byRegister.entries()]
      .filter(([id]) => id !== '__ROOT__' && !registerList.some(r => r.id === id))
      .map(([, v]) => v)
      .sort((a, b) => (a.order ?? 999999) - (b.order ?? 999999)),
  ];

  const directEvents = byRegister.get('__ROOT__')?.events || [];
  let directSort = 1;
  for (const eventType of ['FIELD', 'DISPLAY_FIELD', 'LIST', 'BUTTON']) {
    for (const event of directEvents.filter(e => e.type === eventType)) {
      const itemName = event.label;
      const id = `${rootPageId}/${eventType.toLowerCase()}/${slugify(event.bindField || itemName)}`;
      uniquePush(structure, seen, {
        PAGE_ID: id,
        PARENT_PAGE_ID: rootPageId,
        LEVEL: 1,
        SORT: directSort++,
        TYPE: eventType,
        VIEW_NAME: viewName,
        REGISTER_ID: '',
        ITEM_KEY: event.bindField || itemName,
        DE_NAME: itemName,
        SOURCE_NAME: event.rawLabel || itemName,
        BIND_FIELD: event.bindField || '',
        CONTROL: event.control || '',
        DIRECTORY_PATH: rootSlug,
        URL_PATH: `/de/help/views/${slugify(moduleName)}/${slugify(viewName)}/#${slugify(itemName)}`,
        TITLE: '',
        META_DESCRIPTION: '',
      }, id);
    }
  }

  let tabSort = 1000;
  for (const reg of orderedRegisters) {
    const tabName = `${viewName}-${reg.deName}`;
    const tabId = `${rootPageId}/${reg.id || slugify(tabName)}`;
    const tabSlug = `${rootSlug}/${slugify(reg.id || tabName)}`;
    uniquePush(structure, seen, {
      PAGE_ID: tabId,
      PARENT_PAGE_ID: rootPageId,
      LEVEL: 1,
      SORT: tabSort,
      TYPE: 'REGISTER_TAB',
      VIEW_NAME: viewName,
      REGISTER_ID: reg.id || '',
      ITEM_KEY: reg.id || tabName,
      DE_NAME: tabName,
      SOURCE_NAME: reg.name || '',
      BIND_FIELD: '',
      CONTROL: '',
      DIRECTORY_PATH: tabSlug,
      URL_PATH: `/de/help/views/${slugify(moduleName)}/${slugify(viewName)}/${slugify(reg.id || tabName)}/`,
      TITLE: `ERP Hilfe: ${tabName} in X-ERP`,
      META_DESCRIPTION: `${tabName}: Felder, Listen und Schaltflaechen dieses Registertabs in X-ERP.`,
    }, tabId);
    let itemSort = tabSort + 1;
    const tabEvents = reg.events || [];
    for (const eventType of ['FIELD', 'DISPLAY_FIELD', 'LIST', 'BUTTON']) {
      for (const event of tabEvents.filter(e => e.type === eventType)) {
        const itemName = event.label;
        const id = `${tabId}/${eventType.toLowerCase()}/${slugify(event.bindField || itemName)}`;
        uniquePush(structure, seen, {
          PAGE_ID: id,
          PARENT_PAGE_ID: tabId,
          LEVEL: 2,
          SORT: itemSort++,
          TYPE: eventType,
          VIEW_NAME: viewName,
          REGISTER_ID: reg.id || '',
          ITEM_KEY: event.bindField || itemName,
          DE_NAME: itemName,
          SOURCE_NAME: event.rawLabel || itemName,
          BIND_FIELD: event.bindField || '',
          CONTROL: event.control || '',
          DIRECTORY_PATH: tabSlug,
          URL_PATH: `/de/help/views/${slugify(moduleName)}/${slugify(viewName)}/${slugify(reg.id || tabName)}/#${slugify(itemName)}`,
          TITLE: '',
          META_DESCRIPTION: '',
        }, id);
      }
    }
    tabSort += 1000;
  }

  drafts.push({
    VIEW_NAME: viewName,
    KIND: kind,
    MODULE: moduleName,
    DE_NAME: deName,
    H1: `${deName} in X-ERP`,
    PRIMARY_KEYWORD: 'ERP',
    TITLE: `ERP Hilfe: ${deName} in X-ERP`,
    META_DESCRIPTION: `${deName}: Zweck, Aufbau, Felder und Schaltflaechen der ERP-Ansicht in X-ERP verstaendlich erklaert.`,
    INTRO_DRAFT_DE: textDraft(viewName, kind, deName, moduleName, registerList, counts.FIELD + counts.DISPLAY_FIELD, counts.BUTTON),
    STRUCTURED_DATA_TYPE: 'TechArticle',
    SEO_STATUS: 'ENTWURF_AUS_CODE',
  });
}

const audit = [{
  GENERATED_AT: new Date().toISOString(),
  RAZOR_FILES_READ: allRazorFiles.length,
  RAZOR_BYTES_READ: razorBytesRead,
  VIEW_RAZOR_FILES: razorFiles.length,
  CS_FILES_READ: csFiles.length,
  CS_BYTES_READ: csBytesRead,
  CS_RELEVANT_HITS: csApiClientHits,
  INVENTORY_ROWS: inventory.length,
  STRUCTURE_ROWS: structure.length,
  TEXT_DRAFT_ROWS: drafts.length,
  EDITVIEWS: inventory.filter(r => r.KIND === 'EditView').length,
  WIZARDS: inventory.filter(r => r.KIND === 'Wizard').length,
  LISTVIEWS: inventory.filter(r => r.KIND === 'ListView').length,
}];

writeTable('views-inventory', inventory);
writeTable('views-structure', structure);
writeTable('views-text-drafts', drafts);
writeTable('views-audit', audit);

console.log(JSON.stringify(audit[0], null, 2));
