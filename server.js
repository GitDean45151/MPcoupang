import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import multer from 'multer';
import XLSX from 'xlsx';

// Load environment variables
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

// Resolve static folder paths
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicPath = path.join(__dirname, 'public');

app.use(cors());
app.use(express.json());
app.use(express.static(publicPath));

const isVercel = process.env.VERCEL || process.env.NOW_BUILDER;
const uploadsDir = isVercel ? '/tmp/uploads' : path.join(__dirname, 'uploads');

// Create uploads folder if not exists
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Configure multer for temp file uploads
const upload = multer({ dest: uploadsDir });

// Constants for file paths
const COUPANG_FILE_PATH = path.join(__dirname, 'coupang.xlsx');
const UPDATES_FILE_PATH = path.join(uploadsDir, 'updated_records.json');
const LATEST_SALES_PATH = path.join(uploadsDir, 'latest_sales.xlsx');
const METADATA_PATH = path.join(uploadsDir, 'metadata.json');

// In-memory cache of parsed sheets and joined records
let rawSheets = {
  q1: [],
  q2: [],
  seller: []
};
let joinedRecords = [];

/**
 * Perform Q1 and Q2 Left Joins and calculate commissions
 */
let agencyGrowthRate = 0;
let agencyQoqRate = 0;
let agencyIncentiveRate = 0;
let agencyQ1QoqTotal = 0;
let agencyQ2QoqTotal = 0;
let q1QoqTotalFromFile = 503708883;

// Q1 Group bases for chart comparison
let q1Group1AdSpend = 5506674;
let q1Group2AdSpend = 17493317;
let q1Group3AdSpend = 490090134;
let q1Group4AdSpend = 21605441;

// Q2 Days and reference date tracking based on filename
let q2ElapsedDays = 45;
let q2TotalDays = 91;
let q2ReferenceDate = '260515';

// Baseline tracking variables
let currentQ1QoqTotal = 242727142; // default fallback
let expectedQ1QoqTotal = 0;
let expectedQ2QoqTotal = 0;
let q1VendorLevelData = {}; // vendorId -> { q1AdSpend, q1ColQ }
let mayVendorlistSet = new Set(); // vendorId target list


function performJoinsAndCalculations() {
  console.log('Calculating commissions and joining sheets...');
  
  // Create quick lookup mappings
  const q1Lookup = {};
  rawSheets.q1.forEach(row => {
    const code = String(row.vendorid || '').trim();
    if (code) {
      q1Lookup[code] = {
        q1Revenue: Number(row['이번 분기 광고비']) || 0,
        team: String(row['팀'] || '').trim(),
        marketer: String(row['마케터'] || '').trim()
      };
    }
  });

  const sellerLookup = {};
  rawSheets.seller.forEach(row => {
    const vendorId = String(row.vendorid || '').trim();
    if (vendorId) {
      sellerLookup[vendorId] = String(row['vendor name'] || '').trim();
    }
  });

  // Calculate group-level aggregates from Fact Table (rawSheets.q2)
  let q1_qoq_total = 0;
  let q2_qoq_total = 0;
  
  const processedVendors = rawSheets.q2.map(row => {
    const vendorId = String(row.vendorid || '').trim();
    const commissionGroup = String(row['커미션 그룹'] || '').trim();
    
    const q2Revenue = Number(row['2026_Q2_광고비(~260513)']) || 0;
    const q2AorRevenue = Number(row['2026_Q2_권한설정 광고비']) || 0;
    const q290DayRevenue = Number(row['2026_Q2_90일 커미션대상 광고비']) || 0;
    const q2QoqRevenue = Number(row['2026_Q2_QoQ 대상 권한설정 광고비']) || 0;
    const q1QoqBase = Number(row['직전분기 광고비(D열의 동기간)']) || 0;

    // Accumulate for agency-wide growth: Group 2 and Group 3 QoQ revenues
    if (commissionGroup === 'group 3') {
      q2_qoq_total += q2Revenue; // Use 2Q total spend for Group 3 since they are QoQ-eligible
      q1_qoq_total += q1QoqBase;
    }
    if (commissionGroup === 'group 2') {
      if (q2QoqRevenue > 0) {
        q2_qoq_total += q2QoqRevenue;
        q1_qoq_total += q1QoqBase;
      }
    }
    
    return {
      row,
      vendorId,
      commissionGroup,
      q2Revenue,
      q2AorRevenue,
      q290DayRevenue,
      q2QoqRevenue,
      q1QoqBase
    };
  });

  // Calculate Agency-wide growth and rates (compared actual 2Q QoQ sales directly to actual 1Q QoQ sales)
  agencyGrowthRate = currentQ1QoqTotal > 0 ? (q2_qoq_total / currentQ1QoqTotal - 1) : 0;
  const agency_growth_percent = agencyGrowthRate * 100;
  agencyQ1QoqTotal = currentQ1QoqTotal;
  agencyQ2QoqTotal = q2_qoq_total;
  
  agencyQoqRate = 0;
  agencyIncentiveRate = 0;
  
  if (agencyGrowthRate > 0) {
    if (agency_growth_percent >= 0 && agency_growth_percent < 10) {
      agencyQoqRate = 0.05;
      agencyIncentiveRate = 0.02;
    } else if (agency_growth_percent >= 10 && agency_growth_percent < 20) {
      agencyQoqRate = 0.15;
      agencyIncentiveRate = 0.05;
    } else if (agency_growth_percent >= 20 && agency_growth_percent < 30) {
      agencyQoqRate = 0.175;
      agencyIncentiveRate = 0.07;
    } else if (agency_growth_percent >= 30) {
      agencyQoqRate = 0.20;
      agencyIncentiveRate = 0.10;
    }
  }

  // Calculate dynamic baseline metrics based on the new rules
  if (rawSheets.q2.length === 0) {
    expectedQ1QoqTotal = 0;
    expectedQ2QoqTotal = 0;
    agencyGrowthRate = 0;
    agencyQoqRate = 0;
    agencyIncentiveRate = 0;
    agencyQ1QoqTotal = 0;
    currentQ1QoqTotal = 0;
  } else {
    expectedQ1QoqTotal = q1QoqTotalFromFile;
  }
  
  let expectedQ2Sum = 0;
  processedVendors.forEach(vendor => {
    const vendorId = vendor.vendorId;
    const group = vendor.commissionGroup;
    let vendorQoqBase = 0;
    if (group === 'group 3') {
      vendorQoqBase = vendor.q2Revenue;
    } else if (group === 'group 2') {
      vendorQoqBase = vendor.q2QoqRevenue;
    }
    
    if (mayVendorlistSet.has(vendorId)) {
      expectedQ2Sum += vendorQoqBase * q2TotalDays / q2ElapsedDays;
    } else {
      expectedQ2Sum += vendorQoqBase;
    }
  });
  expectedQ2QoqTotal = Math.round(expectedQ2Sum);

  // Join Fact Table (processedVendors) with Dimensions (LEFT JOIN)
  joinedRecords = processedVendors.map((vendor, index) => {
    const vendorId = vendor.vendorId;
    
    // Left Join dimensions
    const q1Data = q1Lookup[vendorId] || {};
    const vendorName = sellerLookup[vendorId] || '미배정';
    const team = q1Data.team || '미배정';
    const marketer = q1Data.marketer || '미배정';
    const q1Revenue = q1Data.q1Revenue || 0;
    const q2Revenue = vendor.q2Revenue;
    
    // Growth metrics (vendor-level)
    const growthAmount = q2Revenue - q1Revenue;
    const growthRatePercent = q1Revenue > 0 ? (growthAmount / q1Revenue) * 100 : (q2Revenue > 0 ? 100.0 : 0.0);
    
    // 15% new commission: Group 1, Group 2: [2Q 90일 커미션대상 광고비] * 0.15 (rest are 0)
    const isNewCommission = ['group 1', 'group 2'].includes(vendor.commissionGroup);
    const commission15 = isNewCommission ? vendor.q290DayRevenue * 0.15 : 0;
    
    // QoQ Commission eligibility
    const isQoqEligible = ['group 2', 'group 3', 'group 4'].includes(vendor.commissionGroup);
    
    // Expected calculated amounts (using agency-wide rates)
    const expectedQoqCommission = isQoqEligible ? vendor.q2QoqRevenue * agencyQoqRate : 0;

    // Marketer incentives:
    // 1) New commission: 5% of New Target sales (comm90Target * 0.05) paid regardless of growth
    const marketerNewIncentive = isNewCommission ? vendor.q290DayRevenue * 0.05 : 0;
    // 2) QoQ incentive: QoQ Revenue * agencyIncentiveRate
    const marketerQoqIncentive = isQoqEligible ? vendor.q2QoqRevenue * agencyIncentiveRate : 0;
    const expectedMarketerIncentive = marketerNewIncentive + marketerQoqIncentive;

    const isMayVendor = mayVendorlistSet.has(vendorId);
    const q2PredictedRevenue = isMayVendor ? Math.round(q2Revenue * q2TotalDays / q2ElapsedDays) : q2Revenue;

    return {
      id: vendorId || `EXCEL_${index}`,
      vendorId,
      vendorName,
      team,
      marketer,
      commissionGroup: vendor.commissionGroup,
      q1Revenue,
      q2Revenue,
      q2PredictedRevenue,
      growthAmount,
      growthRatePercent,
      isNewCommission,
      commission15, // agency 15% new commission
      isQoqEligible,
      qoqCommissionRate: agencyQoqRate,
      expectedQoqCommission, // agency QoQ commission
      marketerIncentiveRate: agencyIncentiveRate,
      marketerNewIncentive,
      marketerGrowthIncentive: marketerQoqIncentive,
      expectedMarketerIncentive, // marketer total incentive
      isUpdated: false // Reset update status since it's now direct fact table upload
    };
  });

  console.log(`Successfully compiled and calculated ${joinedRecords.length} vendor records.`);
}

/**
 * Normalizes header names to handle variances in spreadsheet imports.
 */
function normalizeHeaderName(h) {
  const norm = String(h || '').trim().toLowerCase().replace(/[\s_-]/g, '');
  if (norm.includes('2026q2광고비') || norm.includes('q2광고비') || norm.includes('q2sales') || norm === '광고비' || norm === '매출' || norm === '매출액' || norm === '실적') {
    return '2026_Q2_광고비(~260513)';
  }
  if (norm.includes('직전분기광고비') || norm.includes('직전분기광고비(d열의동기간)') || norm.includes('직전분기매출')) {
    return '직전분기 광고비(D열의 동기간)';
  }
  if (norm.includes('qoq대상권한설정') || norm.includes('2026q2qoq대상권한설정') || norm.includes('qoq대상광고비')) {
    return '2026_Q2_QoQ 대상 권한설정 광고비';
  }
  if (norm.includes('권한설정광고비') || norm.includes('2026q2권한설정광고비')) {
    return '2026_Q2_권한설정 광고비';
  }
  if (norm.includes('90일커미션') || norm.includes('2026q290일커미션') || norm.includes('90일커미션대상')) {
    return '2026_Q2_90일 커미션대상 광고비';
  }
  if (norm.includes('커미션그룹') || norm === '그룹') {
    return '커미션 그룹';
  }
  if (norm === 'vendorid' || norm === '업체코드' || norm === 'vendor' || norm === 'vendorid코드' || norm === '업체id' || norm === '벤더id') {
    return 'vendorid';
  }
  return h;
}

/**
 * Robust extraction of Vendor ID and Sales values from custom update sheets
 */
function preprocessSalesFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Convert worksheet to 2D array
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  if (rows.length < 2) {
    throw new Error('File has insufficient rows.');
  }

  // 1. Skiprows=1: Skip first row (index 0). Row 2 (index 1) is the candidate header.
  
  // 2. Identify vendorid column by scanning cells for patterns starting with 'A' followed by digits (e.g. A00239521)
  let vendorColIdx = -1;
  let maxMatches = 0;
  for (let c = 0; c < 30; c++) { // scan first 30 columns
    let matches = 0;
    for (let r = 2; r < rows.length; r++) {
      const val = String(rows[r][c] || '').trim();
      if (/^A\d+$/.test(val)) {
        matches++;
      }
    }
    if (matches > maxMatches) {
      maxMatches = matches;
      vendorColIdx = c;
    }
  }

  if (vendorColIdx === -1) {
    // Fallback: look for startsWith 'A00' in data
    for (let c = 0; c < 30; c++) {
      for (let r = 1; r < rows.length; r++) {
        const val = String(rows[r][c] || '').trim();
        if (val.startsWith('A00')) {
          vendorColIdx = c;
          break;
        }
      }
      if (vendorColIdx !== -1) break;
    }
  }

  if (vendorColIdx === -1) {
    throw new Error('벤더 ID 컬럼을 찾을 수 없습니다. (예: A00으로 시작하는 코드)');
  }

  // Find where actual data starts (the first row with a valid vendor ID matching /^A\d+$/)
  let firstVendorRowIdx = -1;
  for (let r = 2; r < rows.length; r++) {
    const val = String(rows[r][vendorColIdx]).trim();
    if (/^A\d+$/.test(val)) {
      firstVendorRowIdx = r;
      break;
    }
  }

  if (firstVendorRowIdx === -1) {
    throw new Error('Coupang 벤더 ID를 포함한 데이터 행이 존재하지 않습니다.');
  }

  // The actual header row for the vendor table is the row right before firstVendorRowIdx
  const rawHeaders = rows[firstVendorRowIdx - 1].map(h => String(h || '').trim());
  rawHeaders[vendorColIdx] = 'vendorid';

  // Normalize header names to guarantee consistency in parsing
  const headers = rawHeaders.map((h, idx) => {
    if (idx === vendorColIdx) return 'vendorid';
    return normalizeHeaderName(h) || `col_${idx}`;
  });

  console.log(`ETL: Detected vendor ID at column ${vendorColIdx}. Data starts at row ${firstVendorRowIdx}. Normalizing headers to:`, headers);

  // Clean data rows
  const cleanRows = [];
  for (let r = firstVendorRowIdx; r < rows.length; r++) {
    const rowVal = rows[r];
    const vendorId = String(rowVal[vendorColIdx] || '').trim();

    // 3. Drop rows where vendorid is null, empty or does not match Coupang ID regex
    if (!vendorId || !/^A\d+$/.test(vendorId)) {
      continue;
    }

    const obj = {};
    headers.forEach((h, idx) => {
      let val = rowVal[idx];

      // 4. Sanitize numeric fields: remove commas and parse as float
      if (typeof val === 'string') {
        const cleanVal = val.replace(/,/g, '').trim();
        if (cleanVal !== '' && !isNaN(cleanVal)) {
          if (h !== 'vendorid' && h !== '최초 권한 설정일' && h !== '권한 해제일') {
            val = parseFloat(cleanVal);
          }
        }
      }
      obj[h] = val;
    });

    cleanRows.push(obj);
  }

  return cleanRows;
}

/**
 * Extracts currentQ1QoqTotal dynamically from the summary section of the Q2 sales file
 */
function extractCurrentQ1QoqTotal(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    let actualColIdx = -1;
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      const row = rows[r];
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c]).trim();
        if (val.includes('% QoQ PA growth (QTD) (actual)')) {
          actualColIdx = c;
          break;
        }
      }
      if (actualColIdx !== -1) {
        if (rows[r + 1]) {
          const rawVal = Number(rows[r + 1][actualColIdx]);
          if (!isNaN(rawVal) && rawVal > 0) {
            return rawVal;
          }
        }
        break;
      }
    }
  } catch (err) {
    console.warn(`Could not extract currentQ1QoqTotal dynamically from ${filePath}:`, err.message);
  }
  return null;
}

/**
 * Parses elapsed/total Q2 days and reference date from filename
 */
function getQ2DaysFromFilename(filename) {
  let elapsedDays = 45; 
  const totalDays = 91; 
  let referenceDate = '260515'; // default fallback for 260515_MP

  const match = filename.match(/(?:20)?(\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/);
  if (match) {
    try {
      const year = 2000 + parseInt(match[1]);
      const month = parseInt(match[2]) - 1; 
      const day = parseInt(match[3]);

      const filenameDate = new Date(year, month, day);
      if (!isNaN(filenameDate.getTime())) {
        const refDate = filenameDate; 
        const q2Start = new Date(year, 3, 1); 
        
        const diffTime = refDate.getTime() - q2Start.getTime();
        const diffDays = Math.floor(diffTime / (24 * 60 * 60 * 1000)) + 1;
        
        if (diffDays >= 1) {
          elapsedDays = Math.min(diffDays, totalDays);
        }

        const yy = String(refDate.getFullYear()).slice(-2);
        const mm = String(refDate.getMonth() + 1).padStart(2, '0');
        const dd = String(refDate.getDate()).padStart(2, '0');
        referenceDate = `${yy}${mm}${dd}`;

        console.log(`Dynamic date parsed from filename "${filename}": ${year}-${month+1}-${day}. Ref date: ${yy}-${mm}-${dd} (${referenceDate}). Q2 Elapsed days: ${elapsedDays}/${totalDays}`);
      }
    } catch (e) {
      console.warn(`Failed to parse date from filename "${filename}":`, e.message);
    }
  } else {
    console.log(`No date pattern matched in filename "${filename}". Using default fallback: ${elapsedDays}/${totalDays}`);
  }
  return { elapsedDays, totalDays, referenceDate };
}

function loadAndInitializeData() {
  console.log('--- Initializing Coupang Ads Dashboard Data ---');
  
  const compiledPath = path.join(__dirname, 'base_data_compiled.json');
  if (!fs.existsSync(compiledPath)) {
    console.error(`CRITICAL ERROR: 'base_data_compiled.json' not found at ${compiledPath}`);
    return;
  }

  try {
    console.log('Loading pre-compiled base data from JSON (Fast-start)...');
    const compiled = JSON.parse(fs.readFileSync(compiledPath, 'utf8'));
    
    // Reconstruct rawSheets.q1
    rawSheets.q1 = Object.entries(compiled.q1Vendors).map(([vendorid, v]) => ({
      vendorid,
      '이번 분기 광고비': v.q1Revenue,
      '팀': v.team,
      '마케터': v.marketer
    }));
    
    // Reconstruct rawSheets.seller
    rawSheets.seller = Object.entries(compiled.q1Vendors).map(([vendorid, v]) => ({
      vendorid,
      'vendor name': v.vendorName
    }));
    
    q1QoqTotalFromFile = compiled.q1QoqTotalFromFile;
    q1Group1AdSpend = compiled.q1GroupBases['group 1'];
    q1Group2AdSpend = compiled.q1GroupBases['group 2'];
    q1Group3AdSpend = compiled.q1GroupBases['group 3'];
    q1Group4AdSpend = compiled.q1GroupBases['group 4'];
    
    mayVendorlistSet = new Set(compiled.mayVendorList);
    
    console.log(`Successfully loaded ${Object.keys(compiled.q1Vendors).length} vendors, Q1 bases, and ${mayVendorlistSet.size} May target vendor IDs.`);
  } catch (error) {
    console.error('Error loading compiled base data JSON:', error);
  }


  // Load Fact Table (either latest uploaded or default 260515_MP.xlsx)
  const latestSalesPath = LATEST_SALES_PATH;
  let factFilePath = COUPANG_FILE_PATH; // fallback
  let factLoaded = false;
  let filenameForDate = '260515_MP.xlsx'; // default fallback

  if (fs.existsSync(latestSalesPath)) {
    factFilePath = latestSalesPath;
    console.log(`Found uploaded sales update at ${latestSalesPath}`);

    // Read original name from metadata.json
    const metadataPath = METADATA_PATH;
    if (fs.existsSync(metadataPath)) {
      try {
        const meta = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
        if (meta.originalName) {
          filenameForDate = meta.originalName;
          console.log(`Resolved original uploaded filename from metadata: ${filenameForDate}`);
        }
      } catch (err) {}
    }

    // Parse Q2 elapsed/total days and reference date dynamically
    const daysInfo = getQ2DaysFromFilename(filenameForDate);
    q2ElapsedDays = daysInfo.elapsedDays;
    q2TotalDays = daysInfo.totalDays;
    q2ReferenceDate = daysInfo.referenceDate;
  } else {
    console.log('No uploaded sales file found. Dashboard initialized with 0 data.');
    rawSheets.q2 = [];
    factFilePath = null;
    q2ElapsedDays = 0;
    q2TotalDays = 91;
    q2ReferenceDate = '적용 없음';
    factLoaded = true;
    currentQ1QoqTotal = 0;
  }

  try {
    if (factFilePath === COUPANG_FILE_PATH) {
      const workbook = XLSX.readFile(COUPANG_FILE_PATH);
      const q2Sheet = workbook.Sheets['2Q'];
      if (q2Sheet) {
        const rows = XLSX.utils.sheet_to_json(q2Sheet, { defval: null });
        rawSheets.q2 = rows.map(row => {
          const vendorid = String(row.vendorid || '').trim();
          return { ...row, vendorid };
        }).filter(row => /^A\d+$/.test(row.vendorid));
        factLoaded = true;
      }
    } else if (factFilePath) {
      console.log(`Preprocessing sales Fact Table from ${factFilePath}...`);
      rawSheets.q2 = preprocessSalesFile(factFilePath);

      // Extract currentQ1QoqTotal dynamically from the summary section of the Q2 sales file
      const dynamicQ1 = extractCurrentQ1QoqTotal(factFilePath);
      if (dynamicQ1 !== null) {
        currentQ1QoqTotal = dynamicQ1;
        console.log(`Found currentQ1QoqTotal dynamically: ${currentQ1QoqTotal}`);
      } else {
        currentQ1QoqTotal = 242727142; // default fallback if we have a file but extraction returned null
      }

      factLoaded = true;
    }
  } catch (err) {
    console.error(`Error loading Fact Table from ${factFilePath}:`, err.message);
  }

  if (factLoaded) {
    performJoinsAndCalculations();
  } else {
    console.error('CRITICAL: Failed to load Fact Table.');
  }
}

// Perform initial data loading
loadAndInitializeData();

// Dynamic uploadsDir handles folder creation above

// ----------------------------------------------------
// API ENDPOINTS
// ----------------------------------------------------

/**
 * GET /api/base-data
 * Returns the cached, fully joined and calculated rows
 */
app.get('/api/base-data', (req, res) => {
  if (rawSheets.q1.length === 0 || rawSheets.seller.length === 0) {
    return res.status(500).json({ error: 'Base data not loaded or failed to parse coupang.xlsx' });
  }
  const latestSalesPath = LATEST_SALES_PATH;
  const hasUpload = fs.existsSync(latestSalesPath);
  res.json({
    records: joinedRecords,
    agencyGrowthRate,
    agencyQoqRate,
    agencyIncentiveRate,
    q1QoqTotal: currentQ1QoqTotal,
    q2QoqTotal: agencyQ2QoqTotal,
    currentQ1QoqTotal,
    currentQ2QoqTotal: agencyQ2QoqTotal,
    expectedQ1QoqTotal,
    expectedQ2QoqTotal,
    lastUpdated: hasUpload ? fs.statSync(latestSalesPath).mtime : null,
    totalUpdatesCount: hasUpload ? 1 : 0,
    q1ChartValues: {
      'group 1': q1Group1AdSpend,
      'group 2': q1Group2AdSpend,
      'group 3': q1Group3AdSpend,
      'group 4': q1Group4AdSpend
    },
    q2ReferenceDate,
    q2ElapsedDays,
    q2TotalDays
  });
});

/**
 * POST /api/upload-update
 * Endpoint to upload a single update file (Excel/CSV)
 */
app.post('/api/upload-update', upload.single('updateFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const filePath = req.file.path;
  const targetPath = LATEST_SALES_PATH;
  
  try {
    console.log(`Processing uploaded sales file: ${req.file.originalname}`);
    const preprocessedRows = preprocessSalesFile(filePath);
    
    // Clean up existing file if any
    if (fs.existsSync(targetPath)) {
      fs.unlinkSync(targetPath);
    }
    
    // Move/rename temp file to latest_sales.xlsx
    fs.renameSync(filePath, targetPath);

    // Save original filename to metadata
    const metadataPath = METADATA_PATH;
    try {
      fs.writeFileSync(metadataPath, JSON.stringify({ originalName: req.file.originalname }), 'utf8');
    } catch (errMeta) {
      console.warn('Failed to save upload metadata:', errMeta.message);
    }

    // Parse dynamic date and elapsed days from uploaded filename
    const daysInfo = getQ2DaysFromFilename(req.file.originalname);
    q2ElapsedDays = daysInfo.elapsedDays;
    q2TotalDays = daysInfo.totalDays;
    q2ReferenceDate = daysInfo.referenceDate;
    
    // Set Fact Table in memory
    rawSheets.q2 = preprocessedRows;
    
    // Extract currentQ1QoqTotal dynamically from the uploaded file
    const dynamicQ1 = extractCurrentQ1QoqTotal(targetPath);
    if (dynamicQ1 !== null) {
      currentQ1QoqTotal = dynamicQ1;
      console.log(`Updated currentQ1QoqTotal dynamically from upload: ${currentQ1QoqTotal}`);
    } else {
      currentQ1QoqTotal = 242727142; // default fallback
    }

    // Perform joins and calculations
    performJoinsAndCalculations();
    
    res.json({
      success: true,
      updatedCount: preprocessedRows.length,
      totalUpdatesCount: 1,
      records: joinedRecords,
      agencyGrowthRate,
      agencyQoqRate,
      agencyIncentiveRate,
      q1QoqTotal: currentQ1QoqTotal,
      q2QoqTotal: agencyQ2QoqTotal,
      currentQ1QoqTotal,
      currentQ2QoqTotal: agencyQ2QoqTotal,
      expectedQ1QoqTotal,
      expectedQ2QoqTotal,
      q1ChartValues: {
        'group 1': q1Group1AdSpend,
        'group 2': q1Group2AdSpend,
        'group 3': q1Group3AdSpend,
        'group 4': q1Group4AdSpend
      },
      q2ReferenceDate,
      q2ElapsedDays,
      q2TotalDays
    });
  } catch (error) {
    console.error('Error processing sales update upload:', error);
    res.status(500).json({ error: `File processing failed: ${error.message}` });
    
    // Clean up temporary upload file
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch (err) {}
  }
});

/**
 * POST /api/reset-updates
 * Clear all uploaded updates and restore original values from coupang.xlsx / 260515_MP.xlsx
 */
app.post('/api/reset-updates', (req, res) => {
  try {
    const latestSalesPath = LATEST_SALES_PATH;
    if (fs.existsSync(latestSalesPath)) {
      fs.unlinkSync(latestSalesPath);
    }
    
    // Reload initial data from disk
    loadAndInitializeData();
    
    res.json({
      success: true,
      records: joinedRecords,
      agencyGrowthRate,
      agencyQoqRate,
      agencyIncentiveRate,
      q1QoqTotal: currentQ1QoqTotal,
      q2QoqTotal: agencyQ2QoqTotal,
      currentQ1QoqTotal,
      currentQ2QoqTotal: agencyQ2QoqTotal,
      expectedQ1QoqTotal,
      expectedQ2QoqTotal,
      q1ChartValues: {
        'group 1': q1Group1AdSpend,
        'group 2': q1Group2AdSpend,
        'group 3': q1Group3AdSpend,
        'group 4': q1Group4AdSpend
      },
      q2ReferenceDate,
      q2ElapsedDays,
      q2TotalDays
    });
  } catch (error) {
    res.status(500).json({ error: `Reset failed: ${error.message}` });
  }
});

// Start Express Server only if not in Vercel serverless environment
if (!process.env.VERCEL) {
  app.listen(port, () => {
    console.log(`Coupang Ads Dashboard backend listening at http://localhost:${port}`);
  });
}

export default app;

