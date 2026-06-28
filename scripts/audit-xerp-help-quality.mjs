import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const workspaceRoot = 'C:/Users/micha/Documents/X-ERP-HELP';
const workbookPath = process.argv[2] || path.join(workspaceRoot, 'X-ERP-HELP.xlsx');
const outDir = path.join(workspaceRoot, 'outputs/help-seo');
const xerpRoot = 'C:/X-ERP/X-ERP';
const resxPath = path.join(xerpRoot, 'X-ERP.Client/ResourceFiles/LzrResource.de-DE.resx');
const registerPath = path.join(xerpRoot, 'X-ERP.Shared/Constants/SystemRegisterData.cs');
const ps = 'powershell.exe';

function decodeXml(s) {
  return String(s ?? '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .trim();
}

function loadTranslations() {
  const text = fs.readFileSync(resxPath, 'utf8');
  const map = new Map();
  const rx = /<data\s+name="([^"]+)"[\s\S]*?<value>([\s\S]*?)<\/value>/g;
  for (const m of text.matchAll(rx)) map.set(m[1], decodeXml(m[2]));
  return map;
}

function loadRegisters() {
  const text = fs.readFileSync(registerPath, 'utf8');
  const byType = new Map();
  const byId = new Map();
  const rx = /new\("([^"]+)",\s*(\d+),\s*"([^"]+)",\s*"([^"]+)"\)/g;
  for (const m of text.matchAll(rx)) {
    const entry = { id: m[1], order: Number(m[2]), english: m[3], type: m[4] };
    byId.set(entry.id, entry);
    const key = entry.type.toLowerCase();
    if (!byType.has(key)) byType.set(key, []);
    byType.get(key).push(entry);
  }
  for (const list of byType.values()) list.sort((a, b) => a.order - b.order);
  return { byType, byId };
}

function runPsJson(script) {
  const output = execFileSync(ps, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], {
    cwd: workspaceRoot,
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  });
  return JSON.parse(output);
}

function loadWorkbookRows() {
  const script = `
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force;
Import-Module ImportExcel;
$pkg=Open-ExcelPackage -Path '${workbookPath.replace(/'/g, "''")}';
try {
  $ws=$pkg.Workbook.Worksheets['de-DE'];
  $rows=@();
  for($r=1;$r -le $ws.Dimension.End.Row;$r++){
    $values=@{};
    for($c=1;$c -le $ws.Dimension.End.Column;$c++){
      $text=[string]$ws.Cells.Item($r,$c).Text;
      if(-not [string]::IsNullOrWhiteSpace($text)){ $values["C$c"]=$text }
    }
    $rows += [pscustomobject]@{
      Row=$r;
      OutlineLevel=[int]$ws.Row($r).OutlineLevel;
      Hidden=[bool]$ws.Row($r).Hidden;
      C1=[string]$ws.Cells.Item($r,1).Text;
      C2=[string]$ws.Cells.Item($r,2).Text;
      C3=[string]$ws.Cells.Item($r,3).Text;
      C4=[string]$ws.Cells.Item($r,4).Text;
      C5=[string]$ws.Cells.Item($r,5).Text;
      C6=[string]$ws.Cells.Item($r,6).Text;
      C10=[string]$ws.Cells.Item($r,10).Text;
      Values=$values
    }
  }
  $rows | ConvertTo-Json -Depth 6
} finally { $pkg.Dispose() }
`;
  return runPsJson(script);
}

function csvEscape(value) {
  const s = String(value ?? '');
  if (/[;"\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function writeCsv(name, rows) {
  fs.mkdirSync(outDir, { recursive: true });
  const headers = Array.from(rows.reduce((set, row) => {
    Object.keys(row).forEach(k => set.add(k));
    return set;
  }, new Set()));
  const csv = [
    headers.map(csvEscape).join(';'),
    ...rows.map(row => headers.map(h => csvEscape(row[h])).join(';')),
  ].join('\r\n');
  fs.writeFileSync(path.join(outDir, name), csv, 'utf8');
}

function norm(s) {
  return String(s ?? '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[„“]/g, '"');
}

function slugPart(s) {
  return String(s ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

const translations = loadTranslations();
const registers = loadRegisters();
const rows = loadWorkbookRows();
const issues = [];
const acceptedTranslationOverrides = new Map([
  ['Partner Contact Person|Partner-Ansprechpartner', 'Anwenderfreundlicher als "Partner Kontaktperson".'],
  ['CRM Group|CRM-Gruppe', 'Deutsche Kompositumschreibung mit Bindestrich.'],
  ['Document Date|Belegdatum', 'Fachlich ueblicher ERP-Begriff.'],
  ['Document Number|Belegnummer', 'Fachlich ueblicher ERP-Begriff.'],
  ['Buyer|Einkäufer', 'Im Belegkontext praeziser als "Kaeufer".'],
  ['Unit Price|Einzelpreis', 'Fachlich ueblicher ERP-Begriff.'],
  ['Warehouse Management|Lagermanagement', 'Als Modul-/Bereichsname freigegeben.'],
  ['Maximum Weight|Maximales Gewicht', 'Natuerlicher als "Maximalgewicht".'],
  ['Production Order|Fertigungsauftrag', 'Im ERP-Kontext akzeptierter Begriff.'],
  ['Search|Suche', 'Als Seiten-/Bereichsname besser als Verb "Suchen".'],
  ['Link|Verknüpfung', 'In der Hilfe sprechender als englisches Lehnwort.'],
  ['Completion|Erledigung', 'Im Prozesskontext sprechender.'],
  ['Finance|Finanzwesen', 'Als Bereichsname freigegeben.'],
  ['Dunning|Mahnwesen', 'Als Bereichsname freigegeben.'],
  ['Report Designer|Reportdesigner', 'Freigegebene Schreibweise als deutscher Fachbegriff.'],
  ['Dashboard Designer|Dashboarddesigner', 'Freigegebene Schreibweise als deutscher Fachbegriff.'],
  ['Country Packages|Länderpakete', 'Kuerzerer Navigationsbegriff.'],
  ['ELSTER Export|ELSTER-Export', 'Deutsche Kompositumschreibung mit Bindestrich.'],
  ['VAT Advance Return|Umsatzsteuer-Voranmeldung', 'Ausgeschriebener Fachbegriff.'],
  ['Information|Auskunft', 'Im Kontext der Ansicht freigegeben.'],
  ['Support Settings|Support-Einstellungen', 'Deutsche Kompositumschreibung mit Bindestrich.'],
]);
const acceptedTranslations = [];

function issue(row, category, severity, message, expected = '', actual = '', evidence = '') {
  issues.push({
    Row: row?.Row ?? '',
    OutlineLevel: row?.OutlineLevel ?? '',
    Category: category,
    Severity: severity,
    Message: message,
    Expected: expected,
    Actual: actual,
    Evidence: evidence,
  });
}

// Translation audit: compare Original Text with current German resource when a resource key exists.
for (const row of rows.slice(1)) {
  const original = norm(row.C2);
  const german = norm(row.C1);
  if (!original || !german) continue;
  if (translations.has(original)) {
    const expected = norm(translations.get(original));
    const overrideKey = `${original}|${german}`;
    if (acceptedTranslationOverrides.has(overrideKey)) {
      acceptedTranslations.push({
        Row: row.Row,
        Original: original,
        ResourceTranslation: expected,
        AcceptedTranslation: german,
        Reason: acceptedTranslationOverrides.get(overrideKey),
      });
      continue;
    }
    if (expected && expected !== german) {
      issue(row, 'TRANSLATION', 'HIGH', 'German label differs from current de-DE resource.', expected, german, `Original Text=${original}`);
    }
  }
}

// Register tab order audit inside Edit/Wizard blocks. Level 2 is view, level 3 register, level 4 fields in current sheet.
let currentView = null;
let currentRegisterType = null;
let seenRegisters = [];
for (const row of rows.slice(1)) {
  if (row.OutlineLevel === 2 && /Edit$|Wizard$/.test(row.C1)) {
    if (currentView) checkRegisterOrder(currentView, currentRegisterType, seenRegisters);
    currentView = row;
    currentRegisterType = row.C1.replace(/Edit$/, '').replace(/Wizard$/, 'Wizard');
    if (row.C1 === 'ArticleEdit') currentRegisterType = 'Article';
    seenRegisters = [];
  } else if (currentView && row.OutlineLevel === 3 && row.C2) {
    const parentBreak = /Edit$|Wizard$/.test(row.C1);
    if (!parentBreak) seenRegisters.push(row);
  }
}
if (currentView) checkRegisterOrder(currentView, currentRegisterType, seenRegisters);

function checkRegisterOrder(viewRow, registerType, registerRows) {
  const expectedList = registers.byType.get(String(registerType).toLowerCase());
  if (!expectedList || expectedList.length === 0 || registerRows.length === 0) return;
  const expectedByEnglish = new Map(expectedList.map((r, idx) => [r.english.toLowerCase(), { ...r, idx }]));
  const expectedByGerman = new Map(expectedList.map((r, idx) => [norm(translations.get(r.english) || r.english).toLowerCase(), { ...r, idx }]));
  const actual = [];
  for (const rr of registerRows) {
    const key1 = norm(rr.C2).toLowerCase();
    const key2 = norm(rr.C1).toLowerCase();
    const match = expectedByEnglish.get(key1) || expectedByGerman.get(key2);
    if (match) actual.push({ row: rr, idx: match.idx, id: match.id, expectedGerman: translations.get(match.english) || match.english, english: match.english });
  }
  for (let i = 1; i < actual.length; i++) {
    if (actual[i].idx < actual[i - 1].idx) {
      issue(
        actual[i].row,
        'ORDER',
        'HIGH',
        `Register tab order differs from SystemRegisterData in ${viewRow.C1}.`,
        `After ${actual[i - 1].expectedGerman}`,
        actual[i].row.C1,
        `Register ${actual[i].id} should not appear after ${actual[i - 1].id}`
      );
    }
  }
  for (const rr of registerRows) {
    const key1 = norm(rr.C2).toLowerCase();
    const key2 = norm(rr.C1).toLowerCase();
    const match = expectedByEnglish.get(key1) || expectedByGerman.get(key2);
    if (match) {
      const expectedGerman = norm(translations.get(match.english) || match.english);
      if (expectedGerman && norm(rr.C1) !== expectedGerman) {
        issue(rr, 'REGISTER_TRANSLATION', 'MEDIUM', 'Register label differs from current resource translation.', expectedGerman, rr.C1, `Register=${match.id}`);
      }
    }
  }
}

// SEO / user usefulness audit for rows that look like pages.
for (const row of rows.slice(1)) {
  if (row.OutlineLevel > 3) continue;
  const topic = norm(row.C1);
  const pathValue = norm(row.C10 || row.C3);
  const isFieldRow = Boolean(norm(row.C6));
  if (!topic) {
    issue(row, 'CONTENT', 'MEDIUM', 'Page/topic row has no German topic text.', 'Non-empty topic', topic);
  }
  if (topic && topic.length < 3 && !isFieldRow && row.OutlineLevel <= 2) {
    issue(row, 'CONTENT', 'LOW', 'Topic is too short to be useful as heading/navigation text.', 'Descriptive heading', topic);
  }
  if (pathValue && /[A-ZÄÖÜ\s]/.test(pathValue)) {
    issue(row, 'SEO', 'MEDIUM', 'URL/path candidate contains uppercase or spaces; slugs should be lowercase and stable.', 'lowercase-hyphen-slug', pathValue);
  }
  if (topic && /^(Edit|View|Wizard)$/i.test(topic)) {
    issue(row, 'CONTENT', 'HIGH', 'Heading is generic and not useful for users or search.', 'Specific page purpose', topic);
  }
}

// Duplicate weak labels within same parent section.
let parentKey = '';
const seenByParent = new Map();
for (const row of rows.slice(1)) {
  if (row.OutlineLevel <= 3) {
    parentKey = `${row.Row}:${row.C1}`;
    seenByParent.set(parentKey, new Map());
    continue;
  }
  if (!parentKey || !row.C1) continue;
  const bucket = seenByParent.get(parentKey);
  const key = norm(row.C1).toLowerCase();
  if (bucket.has(key) && ['artikel', 'name', 'datum', 'position'].includes(key)) {
    issue(row, 'CONTENT', 'LOW', 'Repeated generic field label under one section may need a more specific help text.', 'Specific field description', row.C1, `First seen row ${bucket.get(key)}`);
  } else {
    bucket.set(key, row.Row);
  }
}

const summary = [
  { Metric: 'Rows checked', Value: rows.length },
  { Metric: 'Translation mismatches', Value: issues.filter(i => i.Category === 'TRANSLATION').length },
  { Metric: 'Register translation issues', Value: issues.filter(i => i.Category === 'REGISTER_TRANSLATION').length },
  { Metric: 'Order issues', Value: issues.filter(i => i.Category === 'ORDER').length },
  { Metric: 'SEO/content issues', Value: issues.filter(i => i.Category === 'SEO' || i.Category === 'CONTENT').length },
  { Metric: 'High severity issues', Value: issues.filter(i => i.Severity === 'HIGH').length },
  { Metric: 'Medium severity issues', Value: issues.filter(i => i.Severity === 'MEDIUM').length },
  { Metric: 'Low severity issues', Value: issues.filter(i => i.Severity === 'LOW').length },
];

writeCsv('quality-audit-issues.csv', issues);
writeCsv('quality-audit-summary.csv', summary);
writeCsv('quality-audit-accepted-translations.csv', acceptedTranslations);
fs.writeFileSync(path.join(outDir, 'quality-audit-issues.json'), JSON.stringify(issues, null, 2), 'utf8');
fs.writeFileSync(path.join(outDir, 'quality-audit-summary.json'), JSON.stringify(summary, null, 2), 'utf8');
fs.writeFileSync(path.join(outDir, 'quality-audit-accepted-translations.json'), JSON.stringify(acceptedTranslations, null, 2), 'utf8');

console.log(JSON.stringify({ summary, firstIssues: issues.slice(0, 20) }, null, 2));
