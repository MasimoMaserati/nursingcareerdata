import { writeFile, mkdir } from 'node:fs/promises';
import { unzipSync } from 'fflate';
import * as XLSX from 'xlsx';

const YEAR = 2025;
const OCC_CODE = '29-1141';
const SOURCE_URL = 'https://www.bls.gov/oes/special-requests/oesm25st.zip';
const SOURCE_PAGE = 'https://www.bls.gov/oes/2025/may/oessrcst.htm';

const states = [
  ['Alabama','alabama'],['Alaska','alaska'],['Arizona','arizona'],['Arkansas','arkansas'],
  ['California','california'],['Colorado','colorado'],['Connecticut','connecticut'],['Delaware','delaware'],
  ['District of Columbia','district-of-columbia'],['Florida','florida'],['Georgia','georgia'],['Hawaii','hawaii'],
  ['Idaho','idaho'],['Illinois','illinois'],['Indiana','indiana'],['Iowa','iowa'],['Kansas','kansas'],
  ['Kentucky','kentucky'],['Louisiana','louisiana'],['Maine','maine'],['Maryland','maryland'],
  ['Massachusetts','massachusetts'],['Michigan','michigan'],['Minnesota','minnesota'],['Mississippi','mississippi'],
  ['Missouri','missouri'],['Montana','montana'],['Nebraska','nebraska'],['Nevada','nevada'],
  ['New Hampshire','new-hampshire'],['New Jersey','new-jersey'],['New Mexico','new-mexico'],['New York','new-york'],
  ['North Carolina','north-carolina'],['North Dakota','north-dakota'],['Ohio','ohio'],['Oklahoma','oklahoma'],
  ['Oregon','oregon'],['Pennsylvania','pennsylvania'],['Rhode Island','rhode-island'],
  ['South Carolina','south-carolina'],['South Dakota','south-dakota'],['Tennessee','tennessee'],
  ['Texas','texas'],['Utah','utah'],['Vermont','vermont'],['Virginia','virginia'],
  ['Washington','washington'],['West Virginia','west-virginia'],['Wisconsin','wisconsin'],['Wyoming','wyoming']
];
const allowed = new Map(states);

function numeric(value) {
  if (value === null || value === undefined || value === '' || value === '*' || value === '**' || value === '—') return null;
  const parsed = Number(String(value).replaceAll(',', '').trim());
  return Number.isFinite(parsed) ? parsed : null;
}

console.log(`Downloading BLS OEWS May ${YEAR} state workbook...`);
const response = await fetch(SOURCE_URL, { headers: { 'User-Agent': 'NursingCareerData/1.0 (public wage data ingestion)' } });
if (!response.ok) throw new Error(`BLS download failed: HTTP ${response.status}`);
const zipBytes = new Uint8Array(await response.arrayBuffer());
const archive = unzipSync(zipBytes);
const workbookEntry = Object.entries(archive).find(([name]) => /state.*m2025.*\.xlsx$/i.test(name));
if (!workbookEntry) throw new Error('Could not find the May 2025 state XLSX workbook inside the BLS archive.');

const workbook = XLSX.read(workbookEntry[1], { type: 'array' });
const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: null, raw: false });
if (!rows.length) throw new Error('The BLS state workbook contains no rows.');

const headerNames = Object.keys(rows[0]);
for (const required of ['AREA_TITLE','OCC_CODE','A_MEAN','A_MEDIAN','H_MEAN','TOT_EMP']) {
  if (!headerNames.includes(required)) throw new Error(`Unexpected BLS workbook schema: missing ${required}`);
}

const byState = new Map();
for (const row of rows) {
  if (String(row.OCC_CODE).trim() !== OCC_CODE) continue;
  if (row.NAICS != null && String(row.NAICS).trim() !== '000000') continue;
  if (row.OWN_CODE != null && String(row.OWN_CODE).trim() !== '1235') continue;
  const state = String(row.AREA_TITLE ?? '').trim();
  const stateSlug = allowed.get(state);
  if (!stateSlug) continue;
  if (byState.has(state)) continue;
  byState.set(state, {
    state,
    stateSlug,
    occupation: OCC_CODE,
    occupationTitle: 'Registered Nurses',
    referenceYear: YEAR,
    referencePeriod: 'May',
    annualMeanWage: numeric(row.A_MEAN),
    annualMedianWage: numeric(row.A_MEDIAN),
    hourlyMeanWage: numeric(row.H_MEAN),
    hourlyMedianWage: numeric(row.H_MEDIAN),
    employment: numeric(row.TOT_EMP),
    annual10thPercentile: numeric(row.A_PCT10),
    annual25thPercentile: numeric(row.A_PCT25),
    annual75thPercentile: numeric(row.A_PCT75),
    annual90thPercentile: numeric(row.A_PCT90),
    source: 'U.S. Bureau of Labor Statistics, Occupational Employment and Wage Statistics',
    sourceUrl: SOURCE_PAGE,
    downloadUrl: SOURCE_URL
  });
}

const output = states.map(([state]) => byState.get(state)).filter(Boolean);
if (output.length < 45) {
  throw new Error(`Only ${output.length} of 50 state/DC records matched. Refusing to publish an incomplete or mis-mapped dataset.`);
}
await mkdir(new URL('../src/data/', import.meta.url), { recursive: true });
await writeFile(new URL('../src/data/rn-state-wages.json', import.meta.url), JSON.stringify(output, null, 2) + '\n');
console.log(`Generated ${output.length} verified RN state/DC records for May ${YEAR}.`);
