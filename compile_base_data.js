import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import XLSX from 'xlsx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const COUPANG_FILE_PATH = path.join(__dirname, 'coupang.xlsx');
const Q1_AGENCY_PATH = path.join(__dirname, '2026Q1 Agency commission-MP(PA).xlsx');
const MAY_VENDORLIST_PATH = path.join(__dirname, '2026_Q2_May_vendorlist_MP.xlsx');
const OUTPUT_PATH = path.join(__dirname, 'base_data_compiled.json');

function compile() {
  console.log('Starting base data compilation...');

  if (!fs.existsSync(COUPANG_FILE_PATH)) {
    console.error(`Error: ${COUPANG_FILE_PATH} not found.`);
    process.exit(1);
  }

  // 1. Parse coupang.xlsx
  console.log('Parsing coupang.xlsx...');
  const workbook = XLSX.readFile(COUPANG_FILE_PATH);
  
  const q1Sheet = workbook.Sheets['1Q'];
  let q1Vendors = {};
  if (q1Sheet) {
    const q1Rows = XLSX.utils.sheet_to_json(q1Sheet, { defval: null });
    q1Rows.forEach(row => {
      const vendorid = String(row['업체코드'] || row['vendorid'] || '').trim();
      if (/^A\d+$/.test(vendorid)) {
        q1Vendors[vendorid] = {
          q1Revenue: Number(row['이번 분기 광고비']) || 0,
          team: String(row['팀'] || '').trim(),
          marketer: String(row['마케터'] || '').trim(),
          vendorName: '미배정'
        };
      }
    });
    console.log(`Parsed ${Object.keys(q1Vendors).length} Q1 vendors.`);
  }

  const sellerSheet = workbook.Sheets['2Q seller'];
  let sellerCount = 0;
  if (sellerSheet) {
    const sellerRows = XLSX.utils.sheet_to_json(sellerSheet, { defval: null });
    sellerRows.forEach(row => {
      const vendorid = String(row['vendor id'] || row['vendorid'] || '').trim();
      if (/^A\d+$/.test(vendorid)) {
        sellerCount++;
        const name = String(row['vendor name'] || '').trim();
        if (q1Vendors[vendorid]) {
          q1Vendors[vendorid].vendorName = name;
        } else {
          // If vendor is in 2Q but not in 1Q, initialize it
          q1Vendors[vendorid] = {
            q1Revenue: 0,
            team: '미배정',
            marketer: '미배정',
            vendorName: name
          };
        }
      }
    });
    console.log(`Parsed ${sellerCount} seller mappings.`);
  }

  // 1.5 Parse 2Q sheet to update team and marketer mappings for Q2 vendors
  const q2Sheet = workbook.Sheets['2Q'];
  if (q2Sheet) {
    const q2Rows = XLSX.utils.sheet_to_json(q2Sheet, { defval: null });
    let q2Updated = 0;
    q2Rows.forEach(row => {
      const vendorid = String(row['vendorid'] || row['vendor id'] || '').trim();
      if (/^A\d+$/.test(vendorid)) {
        const team = String(row['팀'] || '').trim();
        const marketer = String(row['마케터'] || '').trim();
        if (q1Vendors[vendorid]) {
          if (team && team !== '미배정' && team !== 'null') {
            q1Vendors[vendorid].team = team;
          }
          if (marketer && marketer !== '미배정' && marketer !== 'null') {
            q1Vendors[vendorid].marketer = marketer;
          }
          q2Updated++;
        }
      }
    });
    console.log(`Updated team/marketer for ${q2Updated} vendors from 2Q sheet.`);
  }

  // 2. Parse 2026Q1 Agency commission-MP(PA).xlsx
  let q1QoqTotalFromFile = 503708883;
  let q1GroupBases = {
    'group 1': 5506674,
    'group 2': 17493317,
    'group 3': 490090134,
    'group 4': 21605441
  };

  if (fs.existsSync(Q1_AGENCY_PATH)) {
    console.log('Parsing Q1 Agency commission file...');
    const q1Wb = XLSX.readFile(Q1_AGENCY_PATH);
    const q1SummarySheet = q1Wb.Sheets['Total summary'];
    if (q1SummarySheet) {
      const q1SummaryRows = XLSX.utils.sheet_to_json(q1SummarySheet);
      let sumQoq = 0;
      let g1New = 0;
      let g2New = 0, g2Qoq = 0;
      let g3Qoq = 0;
      let g4Qoq = 0;

      q1SummaryRows.forEach(row => {
        const group = String(row['커미션 그룹'] || '').trim();
        const val = Number(row['QoQ growth 벤더 이번 분기 광고비']) || 0;
        if (group === 'group 2' || group === 'group 3') {
          sumQoq += val;
        }

        const newVendorsAdSpend = Number(row['90일 이내 신규/재활성화 벤더 광고비']) || 0;
        const qoqAdSpend = Number(row['QoQ growth 수수료 대상 광고비']) || 0;
        
        if (group === 'group 1') {
          g1New = newVendorsAdSpend;
        } else if (group === 'group 2') {
          g2New = newVendorsAdSpend;
          g2Qoq = qoqAdSpend;
        } else if (group === 'group 3') {
          g3Qoq = qoqAdSpend;
        } else if (group === 'group 4') {
          g4Qoq = qoqAdSpend;
        }
      });

      if (sumQoq > 0) q1QoqTotalFromFile = sumQoq;
      if (g1New > 0) q1GroupBases['group 1'] = g1New;
      if (g2New > 0 || g2Qoq > 0) q1GroupBases['group 2'] = g2New + g2Qoq;
      if (g3Qoq > 0) q1GroupBases['group 3'] = g3Qoq;
      if (g4Qoq > 0) q1GroupBases['group 4'] = g4Qoq;

      console.log(`Parsed Q1 Group Bases: G1=${q1GroupBases['group 1']}, G2=${q1GroupBases['group 2']}, G3=${q1GroupBases['group 3']}, G4=${q1GroupBases['group 4']}`);
    }
  }

  // 3. Parse May vendor list
  let mayVendorList = [];
  if (fs.existsSync(MAY_VENDORLIST_PATH)) {
    console.log('Parsing May vendorlist...');
    const mayWb = XLSX.readFile(MAY_VENDORLIST_PATH);
    const maySheet = mayWb.Sheets[mayWb.SheetNames[0]];
    const mayRows = XLSX.utils.sheet_to_json(maySheet, { header: 1, defval: '' });
    const mayHeader = mayRows[1] || [];
    const mayVendorColIdx = mayHeader.indexOf('vendor id') !== -1 ? mayHeader.indexOf('vendor id') : mayHeader.indexOf('vender id');

    if (mayVendorColIdx !== -1) {
      let tempSet = new Set();
      for (let i = 2; i < mayRows.length; i++) {
        const vId = String(mayRows[i][mayVendorColIdx] || '').trim();
        if (/^A\d+$/.test(vId)) {
          tempSet.add(vId);
        }
      }
      mayVendorList = Array.from(tempSet);
      console.log(`Parsed ${mayVendorList.length} May vendor IDs.`);
    }
  }

  // Output compiled JSON
  const compiledData = {
    q1Vendors,
    q1GroupBases,
    q1QoqTotalFromFile,
    mayVendorList
  };

  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(compiledData, null, 2), 'utf8');
  console.log(`Successfully compiled base data and saved to ${OUTPUT_PATH} (size: ${(fs.statSync(OUTPUT_PATH).size / 1024).toFixed(2)} KB).`);
}

compile();
