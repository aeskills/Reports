import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  ResponsiveContainer, BarChart, Bar, LineChart, Line, AreaChart, Area,
  PieChart, Pie, Cell, ScatterChart, Scatter, ComposedChart,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ReferenceLine, ZAxis, Brush,
} from 'recharts';

/* ═══════════════════════════════════════════
   CONSTANTS
   ═══════════════════════════════════════════ */

const CHART_TYPES = [
  { value: 'bar',        label: 'Bar Chart',         icon: '📊', tip: 'Best for comparing values across categories' },
  { value: 'horizontal', label: 'Horizontal Bar',    icon: '📊', tip: 'Ideal when category labels are long' },
  { value: 'grouped',    label: 'Grouped Bar Chart', icon: '📊', tip: 'Compare multiple metrics side by side per category' },
  { value: 'stacked',    label: 'Stacked Bar Chart', icon: '📊', tip: 'Show total and composition at once' },
  { value: 'line',       label: 'Line Chart',        icon: '📈', tip: 'Best for tracking changes over time or ordered sequences' },
  { value: 'area',       label: 'Area Chart',        icon: '📈', tip: 'Like line charts but emphasizes volume beneath the line' },
  { value: 'pie',        label: 'Pie Chart',         icon: '🥧', tip: 'Shows proportions of a whole — best with ≤8 categories' },
  { value: 'donut',      label: 'Donut Chart',       icon: '🍩', tip: 'Like pie but with a clean center for summary stats' },
  { value: 'scatter',    label: 'Scatter Plot',      icon: '🔵', tip: 'Reveals correlations between two numeric fields' },
  { value: 'bubble',     label: 'Bubble Chart',      icon: '🫧', tip: 'Scatter + a 3rd metric controls bubble size' },
  { value: 'composed',   label: 'Composed (Bar+Line)', icon: '📉', tip: 'Overlay bars and lines for multi-metric comparison' },
];

const COLOR_PALETTES = {
  Ocean:      ['#3b82f6','#06b6d4','#8b5cf6','#0ea5e9','#6366f1'],
  Sunset:     ['#f97316','#ef4444','#eab308','#f43f5e','#fb923c'],
  Forest:     ['#10b981','#22c55e','#84cc16','#14b8a6','#059669'],
  Monochrome: ['#94a3b8','#64748b','#cbd5e1','#475569','#e2e8f0'],
  Vibrant:    ['#8b5cf6','#ec4899','#3b82f6','#f97316','#10b981'],
};

const PIE_TYPES = ['pie', 'donut'];

/* ═══════════════════════════════════════════
   FIELD METADATA ENGINE (Feature 2)
   ═══════════════════════════════════════════ */

// Known field metadata for engagement dataset
const KNOWN_FIELDS = {
  categorical: ['Org Name', 'Focus Region', 'Focus Region (India Cohort)', 'State', 'Market Area', 'market_area_name', 'org_name', 'India Cohort', 'focus_region', 'state'],
  numeric: ['Current MAU', 'Prev Year MAU', 'WoW Change', 'MoM MAU', 'YOY MAU', 'WAU', 'Delegations', 'Total FL Ever', 'Current Returning MAU', 'Current New MAU', 'Current Monthly Exports', 'FTA', 'Peak MAU All Time', 'Peak MAU QTR', 'Peak MAU YR'],
  ratio: ['WoW MAU Change%', 'Current by Peak MAU All Time', 'Repeat MAU as % of Last Month\'s MAU', 'Penetration'],
  date: ['as_of_date', 'date', 'Date', 'month', 'Month'],
};

// Positive/negative coloring fields
const DIVERGING_FIELDS = ['WoW Change', 'YOY MAU', 'WoW MAU Change%', 'MoM MAU'];

function detectFieldType(colName, data) {
  // Check known fields first
  const lower = colName.toLowerCase().trim();
  for (const knownName of KNOWN_FIELDS.date) {
    if (lower === knownName.toLowerCase()) return 'date';
  }
  for (const knownName of KNOWN_FIELDS.ratio) {
    if (lower === knownName.toLowerCase() || colName === knownName) return 'ratio';
  }
  for (const knownName of KNOWN_FIELDS.numeric) {
    if (lower === knownName.toLowerCase() || colName === knownName) return 'numeric';
  }
  for (const knownName of KNOWN_FIELDS.categorical) {
    if (lower === knownName.toLowerCase() || colName === knownName) return 'categorical';
  }

  // Heuristic detection from data
  if (!data || data.length === 0) return 'categorical';
  const sample = data.slice(0, 50).map(r => r[colName]).filter(v => v != null && v !== '');
  if (sample.length === 0) return 'categorical';

  // Check if date
  const dateCount = sample.filter(v => !isNaN(Date.parse(String(v))) && String(v).match(/\d{4}|\/|-/)).length;
  if (dateCount > sample.length * 0.7) return 'date';

  // Check if numeric
  const numCount = sample.filter(v => !isNaN(Number(v))).length;
  if (numCount > sample.length * 0.7) {
    // Check if it looks like a percentage or ratio
    const hasPercent = colName.match(/%|percent|ratio|penetration/i);
    const maxVal = Math.max(...sample.map(Number).filter(n => !isNaN(n)));
    if (hasPercent || (maxVal <= 100 && maxVal > 0 && colName.match(/rate|pct|change/i))) return 'ratio';
    return 'numeric';
  }

  // Check unique count — categorical if low cardinality
  const uniq = new Set(sample.map(String));
  if (uniq.size <= 30) return 'categorical';

  return 'categorical';
}

function getFieldMetadata(columns, data) {
  const meta = {};
  for (const col of columns) {
    meta[col] = detectFieldType(col, data);
  }
  return meta;
}

function isDivergingField(colName) {
  return DIVERGING_FIELDS.some(f => colName.toLowerCase().includes(f.toLowerCase()));
}

/* ═══════════════════════════════════════════
   CHART RECOMMENDATION ENGINE (Feature 1)
   ═══════════════════════════════════════════ */

function getChartRecommendations(xField, yFields, fieldMeta, data, allColumns) {
  if (!xField || yFields.length === 0) return [];

  const xType = fieldMeta[xField] || 'categorical';
  const yType = fieldMeta[yFields[0]] || 'numeric';
  const uniqueXCount = new Set(data.slice(0, 500).map(r => r[xField])).size;
  const hasThirdMetric = allColumns.some(c => c !== xField && !yFields.includes(c) && (fieldMeta[c] === 'numeric' || fieldMeta[c] === 'ratio'));
  const isMultiY = yFields.length > 1;
  const avgLabelLen = data.slice(0, 20).reduce((sum, r) => sum + String(r[xField] || '').length, 0) / Math.min(20, data.length);

  const recs = [];

  if (xType === 'categorical' && (yType === 'numeric' || yType === 'ratio')) {
    // Category + Numeric
    if (avgLabelLen > 15) {
      recs.push({ type: 'horizontal', reason: 'Long category labels read better horizontally', rank: 1 });
      recs.push({ type: 'bar', reason: 'Classic vertical bar comparison', rank: 2 });
    } else {
      recs.push({ type: 'bar', reason: 'Best for comparing values across categories', rank: 1 });
      if (avgLabelLen > 10) recs.push({ type: 'horizontal', reason: 'Better readability for medium-length labels', rank: 3 });
    }
    if (uniqueXCount <= 8) {
      recs.push({ type: 'pie', reason: `Only ${uniqueXCount} categories — pie shows proportions well`, rank: 3 });
      recs.push({ type: 'donut', reason: 'Clean donut with summary center', rank: 4 });
    }
    if (isMultiY) {
      recs.push({ type: 'grouped', reason: 'Compare multiple metrics per category', rank: 2 });
      recs.push({ type: 'stacked', reason: 'Show total + composition per category', rank: 4 });
    }
    if (hasThirdMetric) {
      recs.push({ type: 'bubble', reason: 'Add a 3rd metric as bubble size for richer insight', rank: 5 });
    }
    recs.push({ type: 'composed', reason: 'Overlay bars and lines for multi-metric view', rank: 6 });
  } else if (xType === 'date') {
    // Date + Numeric
    recs.push({ type: 'line', reason: 'Best for tracking changes over time', rank: 1 });
    recs.push({ type: 'area', reason: 'Emphasizes volume and trends over time', rank: 2 });
    recs.push({ type: 'bar', reason: 'Compare period totals as bars', rank: 3 });
    if (isMultiY) recs.push({ type: 'composed', reason: 'Mix bars and lines for multi-metric time series', rank: 3 });
  } else if ((xType === 'numeric' || xType === 'ratio') && (yType === 'numeric' || yType === 'ratio')) {
    // Numeric + Numeric
    recs.push({ type: 'scatter', reason: 'Reveals correlations between two metrics', rank: 1 });
    if (hasThirdMetric) recs.push({ type: 'bubble', reason: 'Add bubble size for 3D insight', rank: 2 });
    recs.push({ type: 'line', reason: 'Shows trend if X is ordered', rank: 3 });
  } else if (xType === 'categorical' && yType === 'categorical') {
    recs.push({ type: 'grouped', reason: 'Compare category distributions side by side', rank: 1 });
    recs.push({ type: 'bar', reason: 'Count-based bar comparison', rank: 2 });
  }

  // Always add remaining chart types as non-recommended
  const recTypes = new Set(recs.map(r => r.type));
  CHART_TYPES.forEach(ct => {
    if (!recTypes.has(ct.value)) {
      recs.push({ type: ct.value, reason: ct.tip, rank: 10 });
    }
  });

  return recs.sort((a, b) => a.rank - b.rank);
}

/* ═══════════════════════════════════════════
   PRESET TEMPLATES (Feature 4)
   ═══════════════════════════════════════════ */

const PRESET_TEMPLATES = [
  {
    id: 'cohort-mau',
    name: 'Cohort vs Current MAU',
    description: 'Compare MAU across CBSE, State Govt, PM SHRI, Others',
    icon: '📊',
    xAxis: ['Focus Region (India Cohort)', 'India Cohort', 'focus_region'],
    yAxis: ['Current MAU'],
    chartType: 'grouped',
    topN: 'all',
    filterZeros: true,
    aggregation: 'sum',
  },
  {
    id: 'cohort-wow',
    name: 'Cohort vs WoW Change',
    description: 'Week-over-week change by cohort with ±coloring',
    icon: '📉',
    xAxis: ['Focus Region (India Cohort)', 'India Cohort', 'focus_region'],
    yAxis: ['WoW Change'],
    chartType: 'bar',
    topN: 'all',
    filterZeros: true,
    aggregation: 'avg',
  },
  {
    id: 'top15-mau',
    name: 'Top 15 Orgs by MAU',
    description: 'Horizontal bar — highest engagement organizations',
    icon: '🏆',
    xAxis: ['Org Name', 'org_name'],
    yAxis: ['Current MAU'],
    chartType: 'horizontal',
    topN: '15',
    filterZeros: true,
    aggregation: 'sum',
  },
  {
    id: 'mau-wau-scatter',
    name: 'MAU vs WAU Scatter',
    description: 'Identify high-frequency vs low-frequency orgs',
    icon: '🔵',
    xAxis: ['Current MAU'],
    yAxis: ['WAU'],
    chartType: 'scatter',
    topN: 'all',
    filterZeros: true,
    aggregation: 'sum',
  },
  {
    id: 'cohort-bubble',
    name: 'Cohort MAU + Delegations Bubble',
    description: 'Bubble size = Delegations for delegation density',
    icon: '🫧',
    xAxis: ['Focus Region (India Cohort)', 'India Cohort', 'focus_region'],
    yAxis: ['Current MAU', 'Delegations'],
    chartType: 'bubble',
    topN: 'all',
    filterZeros: true,
    aggregation: 'sum',
  },
];

/* ═══════════════════════════════════════════
   HELPERS
   ═══════════════════════════════════════════ */

function formatAxisTick(value) {
  if (typeof value !== 'number') return value;
  const abs = Math.abs(value);
  if (abs >= 1_000_000_000) return (value / 1_000_000_000).toFixed(1).replace(/\.0$/, '') + 'B';
  if (abs >= 1_000_000) return (value / 1_000_000).toFixed(1).replace(/\.0$/, '') + 'M';
  if (abs >= 1_000) return (value / 1_000).toFixed(1).replace(/\.0$/, '') + 'K';
  if (abs < 1 && abs > 0) return value.toFixed(2);
  return value.toFixed(0);
}

function linearRegression(yValues) {
  const n = yValues.length;
  if (n < 2) return { slope: 0, intercept: 0 };
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
  for (let i = 0; i < n; i++) {
    const y = yValues[i] ?? 0;
    sumX += i; sumY += y; sumXY += i * y; sumX2 += i * i;
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) return { slope: 0, intercept: sumY / n };
  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  return { slope, intercept };
}

function computeStats(values) {
  const nums = values.filter(v => v != null && !isNaN(v)).map(Number);
  if (nums.length === 0) return { min: 0, max: 0, avg: 0, sum: 0, count: 0, trend: 'flat', trendPct: 0 };
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  const sum = nums.reduce((a, b) => a + b, 0);
  const avg = sum / nums.length;
  const { slope } = linearRegression(nums);
  const trendPct = nums[0] !== 0 ? ((nums[nums.length - 1] - nums[0]) / Math.abs(nums[0])) * 100 : 0;
  let trend = 'flat';
  if (slope > 0.02) trend = 'upward';
  else if (slope < -0.02) trend = 'downward';
  return { min, max, avg, sum, count: nums.length, trend, trendPct };
}

function generateInsight(yColumns, data) {
  const lines = [];
  for (const col of yColumns) {
    const values = data.map(d => d[col]);
    const stats = computeStats(values);
    const fmtNum = n => {
      if (Math.abs(n) >= 1000) return (n / 1000).toFixed(1) + 'K';
      return Number(n).toFixed(1);
    };
    const trendWord = stats.trend === 'upward' ? 'upward' : stats.trend === 'downward' ? 'downward' : 'stable';
    const trendDesc = stats.trend === 'flat'
      ? 'The trend appears stable with no significant directional movement.'
      : `The trendline indicates a ${Math.abs(stats.trendPct).toFixed(1)}% ${trendWord} trajectory.`;
    lines.push(`${col} averaged ${fmtNum(stats.avg)}, peaking at ${fmtNum(stats.max)} and bottoming at ${fmtNum(stats.min)}. ${trendDesc}`);
  }
  return lines.join(' ');
}

function addTrendData(data, yColumns) {
  return data.map((row, i) => {
    const newRow = { ...row };
    for (const col of yColumns) {
      const values = data.map(d => Number(d[col]) || 0);
      const { slope, intercept } = linearRegression(values);
      newRow[`__trend_${col}`] = slope * i + intercept;
    }
    return newRow;
  });
}

function generateSampleCSV() {
  const orgs = [
    'Kendriya Vidyalaya No.1 Delhi', 'Delhi Public School Dwarka', 'DAV Public School Noida', 
    'Ryan International School', 'Amity International School', 'Army Public School', 
    'State Govt School Jaipur', 'PM SHRI School Bhopal', 'CBSE Academy Delhi', 
    'Podar International School Mumbai', 'Modern School Barakhamba', 'The Heritage School'
  ];
  
  // Weekly dates starting on Fridays
  const dates = [
    '2026-05-01', '2026-05-08', '2026-05-15', '2026-05-22', 
    '2026-05-29', '2026-06-05', '2026-06-12', '2026-06-19'
  ];
  
  const rows = [];
  for (const org of orgs) {
    dates.forEach((date, dateIdx) => {
      // Build a flat array representing columns A to AO (and some extra)
      const colArray = Array(45).fill('');
      
      // Basic info
      colArray[0] = `SCH-${1000 + Math.floor(Math.random() * 9000)}`; // Column A
      colArray[1] = org; // Column B (we will map to "Org Name")
      colArray[2] = date; // Column C (we will map to "as_of_date")
      colArray[3] = dateIdx % 2 === 0 ? 'North' : 'West'; // Column D: Focus Region
      colArray[4] = org.includes('Govt') ? 'State Govt' : org.includes('SHRI') ? 'PM SHRI' : 'CBSE'; // Column E: India Cohort
      
      // M (index 12): Current MAU
      const baseMAU = 500 + Math.floor(Math.random() * 1500) + (dateIdx * 80);
      colArray[12] = baseMAU;
      
      // AA (index 26): Current Returning MAU
      colArray[26] = Math.floor(baseMAU * (0.55 + Math.random() * 0.15));
      
      // AH (index 33): Current New MAU
      colArray[33] = Math.floor(baseMAU * (0.2 + Math.random() * 0.1));
      
      // AO (index 40): Current Monthly Exports
      colArray[40] = Math.floor(baseMAU * (0.04 + Math.random() * 0.03));
      
      // Let's add other columns to be complete
      colArray[11] = Math.floor(baseMAU * 0.45); // Column L: WAU
      colArray[13] = 100 + Math.floor(Math.random() * 300); // Column N: Delegations
      
      // Convert to row object
      const rowObj = {};
      colArray.forEach((val, idx) => {
        let name = `Col_${idx + 1}`;
        if (idx === 1) name = 'Org Name';
        else if (idx === 2) name = 'as_of_date';
        else if (idx === 12) name = 'Current MAU';
        else if (idx === 26) name = 'Current Returning MAU';
        else if (idx === 33) name = 'Current New MAU';
        else if (idx === 40) name = 'Current Monthly Exports';
        
        rowObj[name] = val;
      });
      rows.push(rowObj);
    });
  }
  
  return Papa.unparse(rows);
}

/* ── Aggregation helper (Feature 3) ── */
function aggregateData(data, xAxis, yAxis, mode) {
  if (mode === 'sum' || mode === 'avg' || mode === 'count') {
    const groups = {};
    for (const row of data) {
      const key = String(row[xAxis]);
      if (!groups[key]) {
        groups[key] = { [xAxis]: row[xAxis], __count: 0, __orgName: row['Org Name'] || row['org_name'] || '' };
        yAxis.forEach(col => { groups[key][col] = 0; });
      }
      groups[key].__count++;
      yAxis.forEach(col => { groups[key][col] += Number(row[col]) || 0; });
    }
    const result = Object.values(groups);
    if (mode === 'avg') {
      result.forEach(row => {
        yAxis.forEach(col => { row[col] = row.__count > 0 ? row[col] / row.__count : 0; });
      });
    }
    if (mode === 'count') {
      result.forEach(row => {
        yAxis.forEach(col => { row[col] = row.__count; });
      });
    }
    return result;
  }
  return data;
}

function processChartData(data, config, fieldMeta) {
  const { xAxis, yAxis, chartType, filterZeros, topN, aggregation, bubbleField } = config;
  if (!data || !xAxis || yAxis.length === 0) return [];

  // Step 1: Aggregate data
  let d = aggregateData(data, xAxis, yAxis, aggregation || 'sum');

  const xType = fieldMeta[xAxis] || 'categorical';
  const isNumericX = xType === 'numeric' || xType === 'ratio';

  // Step 2: Build chart points
  d = d.map(row => {
    const point = { 
      [xAxis]: isNumericX ? (row[xAxis] != null ? Number(row[xAxis]) : null) : row[xAxis], 
      __orgName: row['Org Name'] || row['org_name'] || row[xAxis] || '' 
    };
    yAxis.forEach(col => { point[col] = row[col] != null ? Number(row[col]) : null; });
    if (bubbleField) point[bubbleField] = row[bubbleField] != null ? Number(row[bubbleField]) : 1;
    return point;
  });

  // Step 3: Filter null/empty X values
  d = d.filter(row => row[xAxis] != null && row[xAxis] !== '');

  // Step 4: Filter zeros
  if (filterZeros) {
    d = d.filter(row => yAxis.some(col => row[col] != null && row[col] !== 0));
  }

  // Step 5: Limit / Sort data
  if (topN && topN !== 'all' && !PIE_TYPES.includes(chartType)) {
    const n = parseInt(topN, 10);
    if (xType === 'categorical') {
      d = [...d].sort((a, b) => {
        const aVal = yAxis.reduce((sum, col) => sum + Math.abs(Number(a[col]) || 0), 0);
        const bVal = yAxis.reduce((sum, col) => sum + Math.abs(Number(b[col]) || 0), 0);
        return bVal - aVal;
      }).slice(0, n);
    } else {
      d = [...d].sort((a, b) => {
        const aVal = a[xAxis];
        const bVal = b[xAxis];
        if (xType === 'date') {
          const da = new Date(aVal);
          const db = new Date(bVal);
          if (isNaN(da) || isNaN(db)) return String(aVal).localeCompare(String(bVal));
          return da - db;
        }
        return (Number(aVal) || 0) - (Number(bVal) || 0);
      }).slice(0, n);
    }
  } else {
    if (xType === 'date' || xType === 'numeric' || xType === 'ratio') {
      d = [...d].sort((a, b) => {
        const aVal = a[xAxis];
        const bVal = b[xAxis];
        if (xType === 'date') {
          const da = new Date(aVal);
          const db = new Date(bVal);
          if (isNaN(da) || isNaN(db)) return String(aVal).localeCompare(String(bVal));
          return da - db;
        }
        return (Number(aVal) || 0) - (Number(bVal) || 0);
      });
    }
  }

  return d;
}

// Dynamic Column Mapping Helpers
function getMetricColumn(row, colLetter) {
  if (!row) return null;
  const keys = Object.keys(row);
  
  // Map letter to 0-based index
  const indexMap = {
    'M': 12,
    'AA': 26,
    'AH': 33,
    'AO': 40
  };
  const targetIdx = indexMap[colLetter];
  
  // If we can find the key by index
  if (targetIdx !== undefined && targetIdx < keys.length) {
    return keys[targetIdx];
  }
  
  // Heuristic search by name if index is out of bounds or keys are different
  const nameMap = {
    'M': ['current mau', 'mau', 'current_mau', 'col_13'],
    'AA': ['current returning mau', 'rmau', 'returning mau', 'returning_mau', 'current_returning_mau', 'col_27'],
    'AH': ['current new mau', 'new mau', 'new_mau', 'current_new_mau', 'col_34'],
    'AO': ['current monthly exports', 'export', 'exports', 'monthly exports', 'current_monthly_exports', 'col_41']
  };
  
  const searchNames = nameMap[colLetter] || [];
  for (const key of keys) {
    const lowerKey = key.toLowerCase().trim();
    if (searchNames.includes(lowerKey) || searchNames.some(name => lowerKey.includes(name))) {
      return key;
    }
  }
  
  return null;
}

function getOrgNameColumn(row) {
  if (!row) return null;
  const keys = Object.keys(row);
  const candidates = ['org name', 'org_name', 'organization name', 'school name', 'school_name', 'orgname', 'school'];
  for (const key of keys) {
    const lowerKey = key.toLowerCase().trim();
    if (candidates.includes(lowerKey) || candidates.some(cand => lowerKey.includes(cand))) {
      return key;
    }
  }
  return keys[0] || null;
}

function getAsOfDateColumn(row) {
  if (!row) return null;
  const keys = Object.keys(row);
  const candidates = ['as_of_date', 'date', 'as of date', 'asofdate', 'week', 'date_week', 'month'];
  for (const key of keys) {
    const lowerKey = key.toLowerCase().trim();
    if (candidates.includes(lowerKey) || candidates.some(cand => lowerKey.includes(cand))) {
      return key;
    }
  }
  return keys[2] || keys[0] || null;
}

/* ═══════════════════════════════════════════
   CUSTOM TOOLTIP (Feature 3 — Enhanced)
   ═══════════════════════════════════════════ */

const CustomTooltip = ({ active, payload, label, dark }) => {
  if (!active || !payload?.length) return null;
  return (
    <div className={`rounded-lg px-4 py-3 shadow-xl border text-sm max-w-xs ${dark ? 'bg-slate-800 border-slate-700 text-slate-200' : 'bg-white border-slate-200 text-slate-700'}`}>
      <p className="font-semibold mb-1.5 font-mono-data truncate">{label}</p>
      {payload.filter(p => !p.dataKey?.startsWith('__trend_')).map((p, i) => (
        <p key={i} className="flex items-center gap-2">
          <span className="w-2.5 h-2.5 rounded-full inline-block flex-shrink-0" style={{ backgroundColor: p.color }} />
          <span className="opacity-70 truncate">{p.name}:</span>
          <span className={`font-semibold font-mono-data ${typeof p.value === 'number' && p.value < 0 ? 'text-red-400' : ''}`}>
            {typeof p.value === 'number' ? p.value.toLocaleString(undefined, { maximumFractionDigits: 1 }) : p.value}
          </span>
        </p>
      ))}
    </div>
  );
};

/* ═══════════════════════════════════════════
   SEARCHABLE DROPDOWN
   ═══════════════════════════════════════════ */

const SearchableDropdown = ({ options, value, onChange, placeholder, dark, multiple, maxSelections }) => {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState('');
  const ref = useRef(null);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const filtered = options.filter(o => o.toLowerCase().includes(search.toLowerCase()));
  const isSelected = (opt) => multiple ? (value || []).includes(opt) : value === opt;

  const handleSelect = (opt) => {
    if (multiple) {
      const arr = value || [];
      if (arr.includes(opt)) {
        onChange(arr.filter(v => v !== opt));
      } else if (!maxSelections || arr.length < maxSelections) {
        onChange([...arr, opt]);
      }
    } else {
      onChange(opt);
      setOpen(false);
    }
  };

  const displayVal = multiple
    ? (value && value.length > 0 ? value.join(', ') : '')
    : (value || '');

  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className={`w-full text-left px-3 py-2.5 rounded-lg border text-sm flex items-center justify-between transition-all duration-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 ${dark ? 'bg-slate-800/80 border-slate-600 text-slate-200 hover:border-slate-500' : 'bg-white border-slate-300 text-slate-700 hover:border-slate-400'}`}
      >
        <span className={`truncate ${!displayVal ? 'opacity-50' : ''}`}>
          {displayVal || placeholder}
        </span>
        <svg className={`w-4 h-4 transition-transform ${open ? 'rotate-180' : ''} ${dark ? 'text-slate-400' : 'text-slate-500'}`} fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg>
      </button>
      {open && (
        <div className={`absolute z-50 mt-1 w-full rounded-lg border shadow-2xl max-h-56 overflow-hidden ${dark ? 'bg-slate-800 border-slate-600' : 'bg-white border-slate-200'}`}>
          {options.length > 10 && (
            <div className="p-2 border-b border-slate-600/30">
              <input
                type="text"
                value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Search…"
                className={`w-full px-2.5 py-1.5 rounded text-sm focus:outline-none ${dark ? 'bg-slate-700 text-slate-200 placeholder-slate-500' : 'bg-slate-100 text-slate-700 placeholder-slate-400'}`}
              />
            </div>
          )}
          <div className="overflow-y-auto max-h-48">
            {filtered.map(opt => (
              <button
                key={opt}
                type="button"
                onClick={() => handleSelect(opt)}
                className={`w-full text-left px-3 py-2 text-sm flex items-center gap-2 transition-colors ${isSelected(opt) ? (dark ? 'bg-blue-600/20 text-blue-400' : 'bg-blue-50 text-blue-600') : (dark ? 'text-slate-300 hover:bg-slate-700' : 'text-slate-600 hover:bg-slate-100')}`}
              >
                {multiple && (
                  <span className={`w-4 h-4 rounded border flex-shrink-0 flex items-center justify-center text-xs ${isSelected(opt) ? 'bg-blue-500 border-blue-500 text-white' : (dark ? 'border-slate-500' : 'border-slate-300')}`}>
                    {isSelected(opt) && '✓'}
                  </span>
                )}
                {opt}
              </button>
            ))}
            {filtered.length === 0 && (
              <p className={`px-3 py-2 text-sm ${dark ? 'text-slate-500' : 'text-slate-400'}`}>No matches</p>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

/* ═══════════════════════════════════════════
   CHART RENDERER (Feature 3 — Enhanced)
   ═══════════════════════════════════════════ */

const ChartRenderer = ({ config, data, fieldMeta = {}, height = 480, dark, animate = true, onBrushChange }) => {
  const { chartType, xAxis, yAxis, palette, showTrendline, legendLabels, logScale, bubbleField, brushStartIndex, brushEndIndex } = config;
  const colors = COLOR_PALETTES[palette] || COLOR_PALETTES.Ocean;
  const gridStroke = dark ? 'rgba(30,41,59,0.6)' : 'rgba(226,232,240,0.7)';
  const axisStroke = dark ? '#334155' : '#cbd5e1';
  const textFill = dark ? '#94a3b8' : '#64748b';

  // Check if Y fields have diverging (positive/negative) data
  const hasDiverging = yAxis.some(isDivergingField);

  // Process data using our helper
  const chartData = useMemo(() => {
    let d = processChartData(data, config, fieldMeta);
    if (showTrendline && !PIE_TYPES.includes(chartType)) {
      d = addTrendData(d, yAxis);
    }
    return d;
  }, [data, config, fieldMeta, showTrendline, chartType, yAxis]);

  const xType = fieldMeta[xAxis] || 'categorical';
  const isNumericX = xType === 'numeric' || xType === 'ratio';

  // Smart X-axis
  const dataLen = chartData.length;
  const xInterval = dataLen <= 15 ? 0 : dataLen <= 30 ? 1 : dataLen <= 60 ? 2 : Math.floor(dataLen / 20);

  const customTicks = useMemo(() => {
    if (isNumericX) return undefined;
    if (chartData.length <= 12) return undefined;
    const ticks = [];
    const step = Math.ceil(chartData.length / 10);
    for (let i = 0; i < chartData.length; i += step) {
      ticks.push(chartData[i][xAxis]);
    }
    const lastVal = chartData[chartData.length - 1][xAxis];
    if (!ticks.includes(lastVal)) {
      ticks.push(lastVal);
    }
    return ticks;
  }, [chartData, xAxis, isNumericX]);

  const xAxisProps = {
    dataKey: xAxis,
    tick: { fill: textFill, fontSize: 11, fontFamily: "'DM Mono', monospace" },
    stroke: axisStroke,
    angle: (!isNumericX && dataLen > 8) ? -45 : 0,
    textAnchor: (!isNumericX && dataLen > 8) ? 'end' : 'middle',
    height: (!isNumericX && dataLen > 8) ? 80 : 50,
    interval: isNumericX ? undefined : (customTicks ? 0 : xInterval),
    ticks: customTicks,
    tickMargin: 8,
    type: isNumericX ? 'number' : 'category',
    domain: isNumericX ? ['auto', 'auto'] : undefined,
  };

  const yAxisProps = {
    tick: { fill: textFill, fontSize: 11, fontFamily: "'DM Mono', monospace" },
    stroke: axisStroke,
    tickFormatter: formatAxisTick,
    width: 65,
    tickMargin: 4,
    ...(logScale ? { scale: 'log', domain: ['auto', 'auto'], allowDataOverflow: true } : {}),
  };
  const gridProps = { strokeDasharray: '3 6', stroke: gridStroke, vertical: false };
  const legendProps = { wrapperStyle: { color: textFill, fontSize: 12, paddingTop: 8 }, iconType: 'circle', iconSize: 8 };
  const chartMargin = { top: 20, right: 30, bottom: 20, left: 10 };

  const getLabel = (col) => legendLabels?.[col] || col;

  // Positive/negative bar colors
  const getDivergingColor = (value) => {
    if (typeof value !== 'number') return colors[0];
    return value >= 0 ? '#10b981' : '#ef4444';
  };

  const renderBrush = () => {
    if (PIE_TYPES.includes(chartType) || chartType === 'horizontal' || chartType === 'scatter' || chartType === 'bubble' || dataLen <= 25) {
      return null;
    }
    return (
      <Brush
        dataKey={xAxis}
        height={22}
        stroke={dark ? '#475569' : '#cbd5e1'}
        fill={dark ? '#1e293b' : '#f8fafc'}
        tickFormatter={(val) => {
          const s = String(val);
          return s.length > 10 ? s.substring(0, 10) + '…' : s;
        }}
        startIndex={brushStartIndex !== undefined ? brushStartIndex : 0}
        endIndex={brushEndIndex !== undefined ? brushEndIndex : Math.min(dataLen - 1, 24)}
        onChange={onBrushChange}
      />
    );
  };

  // Empty state
  if (chartData.length === 0) {
    return (
      <div className={`flex flex-col items-center justify-center py-16 ${dark ? 'text-slate-500' : 'text-slate-400'}`}>
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" className="mb-3 opacity-50">
          <path d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
        </svg>
        <p className="text-sm font-medium">No data to display</p>
        <p className="text-xs mt-1">Try adjusting filters or selecting different columns</p>
      </div>
    );
  }

  // ── PIE / DONUT ──
  if (PIE_TYPES.includes(chartType)) {
    const pieData = chartData.map(row => ({
      name: String(row[xAxis]),
      value: Math.abs(Number(row[yAxis[0]]) || 0),
    })).filter(d => d.value > 0);
    const total = pieData.reduce((s, d) => s + d.value, 0);
    const innerR = chartType === 'donut' ? 75 : 0;
    return (
      <ResponsiveContainer width="100%" height={height}>
        <PieChart>
          <Pie data={pieData} cx="50%" cy="48%" innerRadius={innerR} outerRadius={height > 200 ? 150 : 60}
            paddingAngle={2} dataKey="value" isAnimationActive={animate} animationDuration={800}
            label={({ name, value }) => `${name}: ${((value / total) * 100).toFixed(1)}%`}
            labelLine={{ stroke: axisStroke, strokeWidth: 1 }}>
            {pieData.map((_, i) => <Cell key={i} fill={colors[i % colors.length]} />)}
          </Pie>
          <Tooltip content={<CustomTooltip dark={dark} />} />
          <Legend {...legendProps} />
        </PieChart>
      </ResponsiveContainer>
    );
  }

  // ── SCATTER ──
  if (chartType === 'scatter') {
    return (
      <ResponsiveContainer width="100%" height={height}>
        <ScatterChart margin={chartMargin}>
          <CartesianGrid {...gridProps} />
          <XAxis {...xAxisProps} name={xAxis} />
          <YAxis {...yAxisProps} name={yAxis[0]} />
          <Tooltip content={<CustomTooltip dark={dark} />} />
          <Legend {...legendProps} />
          <Scatter name={`${xAxis} vs ${yAxis[0]}`} data={chartData} fill={colors[0]} isAnimationActive={animate}>
            {chartData.map((_, j) => <Cell key={j} fill={colors[0]} />)}
          </Scatter>
        </ScatterChart>
      </ResponsiveContainer>
    );
  }

  // ── BUBBLE ──
  if (chartType === 'bubble') {
    const sizeField = bubbleField || yAxis[1] || yAxis[0];
    return (
      <ResponsiveContainer width="100%" height={height}>
        <ScatterChart margin={chartMargin}>
          <CartesianGrid {...gridProps} />
          <XAxis {...xAxisProps} name={xAxis} />
          <YAxis {...yAxisProps} dataKey={yAxis[0]} name={yAxis[0]} />
          <ZAxis dataKey={sizeField} range={[40, 600]} name={sizeField} />
          <Tooltip content={<CustomTooltip dark={dark} />} />
          <Legend {...legendProps} />
          <Scatter name={`${yAxis[0]} (size: ${sizeField})`} data={chartData} fill={colors[0]} isAnimationActive={animate}>
            {chartData.map((_, j) => <Cell key={j} fill={colors[j % colors.length]} opacity={0.7} />)}
          </Scatter>
        </ScatterChart>
      </ResponsiveContainer>
    );
  }

  // ── HORIZONTAL BAR ──
  if (chartType === 'horizontal') {
    const bLen = chartData.length;
    const barH = bLen <= 6 ? 44 : bLen <= 12 ? 32 : bLen <= 24 ? 22 : 14;
    return (
      <ResponsiveContainer width="100%" height={Math.max(height, bLen * (barH + 10))}>
        <BarChart data={chartData} layout="vertical" margin={{ ...chartMargin, left: 100 }} barCategoryGap="12%">
          <CartesianGrid {...gridProps} horizontal={false} vertical={true} />
          <XAxis type="number" tick={{ fill: textFill, fontSize: 11, fontFamily: "'DM Mono', monospace" }} stroke={axisStroke} tickFormatter={formatAxisTick} />
          <YAxis type="category" dataKey={xAxis} tick={{ fill: textFill, fontSize: 11, fontFamily: "'DM Mono', monospace" }} stroke={axisStroke} width={95} tickMargin={4} interval={0} />
          <Tooltip content={<CustomTooltip dark={dark} />} cursor={{ fill: dark ? 'rgba(59,130,246,0.06)' : 'rgba(59,130,246,0.04)' }} />
          <Legend {...legendProps} />
          {yAxis.map((col, i) => (
            hasDiverging && isDivergingField(col)
              ? <Bar key={col} dataKey={col} name={getLabel(col)} barSize={barH} minPointSize={3}
                  isAnimationActive={animate} animationDuration={800} radius={[0,6,6,0]}>
                  {chartData.map((entry, j) => <Cell key={j} fill={getDivergingColor(entry[col])} />)}
                </Bar>
              : <Bar key={col} dataKey={col} name={getLabel(col)} fill={colors[i % colors.length]}
                  barSize={barH} minPointSize={3} isAnimationActive={animate} animationDuration={800} radius={[0,6,6,0]} />
          ))}
        </BarChart>
      </ResponsiveContainer>
    );
  }

  // ── COMPOSED ──
  if (chartType === 'composed') {
    return (
      <ResponsiveContainer width="100%" height={height}>
        <ComposedChart data={chartData} margin={chartMargin}>
          <CartesianGrid {...gridProps} />
          <XAxis {...xAxisProps} />
          <YAxis {...yAxisProps} />
          <Tooltip content={<CustomTooltip dark={dark} />} />
          <Legend {...legendProps} />
          {yAxis.map((col, i) =>
            i === 0
              ? <Bar key={col} dataKey={col} name={getLabel(col)} fill={colors[i % colors.length]} radius={[6,6,0,0]} barSize={32} isAnimationActive={animate} animationDuration={800} />
              : <Line key={col} dataKey={col} name={getLabel(col)} stroke={colors[i % colors.length]} strokeWidth={2.5} dot={{ r: 3, fill: colors[i % colors.length] }} isAnimationActive={animate} animationDuration={800} />
          )}
          {showTrendline && yAxis.map((col, i) => (
            <Line key={`trend_${col}`} dataKey={`__trend_${col}`} name={`Trend (${getLabel(col)})`}
              stroke={colors[i % colors.length]} strokeWidth={1.5} strokeDasharray="6 3" dot={false} isAnimationActive={animate} />
          ))}
          {renderBrush()}
        </ComposedChart>
      </ResponsiveContainer>
    );
  }

  // ── AREA ──
  if (chartType === 'area') {
    return (
      <ResponsiveContainer width="100%" height={height}>
        <AreaChart data={chartData} margin={chartMargin}>
          <defs>
            {yAxis.map((col, i) => (
              <linearGradient key={col} id={`grad_${col}`} x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={colors[i % colors.length]} stopOpacity={0.4} />
                <stop offset="95%" stopColor={colors[i % colors.length]} stopOpacity={0.02} />
              </linearGradient>
            ))}
          </defs>
          <CartesianGrid {...gridProps} />
          <XAxis {...xAxisProps} />
          <YAxis {...yAxisProps} />
          <Tooltip content={<CustomTooltip dark={dark} />} />
          <Legend {...legendProps} />
          {yAxis.map((col, i) => (
            <Area key={col} type="monotone" dataKey={col} name={getLabel(col)}
              stroke={colors[i % colors.length]} fill={`url(#grad_${col})`} strokeWidth={2.5}
              dot={{ r: 2, fill: colors[i % colors.length], strokeWidth: 0 }}
              isAnimationActive={animate} animationDuration={800} />
          ))}
          {showTrendline && yAxis.map((col, i) => (
            <Line key={`trend_${col}`} dataKey={`__trend_${col}`} name={`Trend (${getLabel(col)})`}
              stroke={colors[i % colors.length]} strokeWidth={1.5} strokeDasharray="6 3" dot={false} isAnimationActive={animate} />
          ))}
          {renderBrush()}
        </AreaChart>
      </ResponsiveContainer>
    );
  }

  // ── LINE ──
  if (chartType === 'line') {
    return (
      <ResponsiveContainer width="100%" height={height}>
        <LineChart data={chartData} margin={chartMargin}>
          <CartesianGrid {...gridProps} />
          <XAxis {...xAxisProps} />
          <YAxis {...yAxisProps} />
          <Tooltip content={<CustomTooltip dark={dark} />} />
          <Legend {...legendProps} />
          {yAxis.map((col, i) => (
            <Line key={col} type="monotone" dataKey={col} name={getLabel(col)}
              stroke={colors[i % colors.length]} strokeWidth={2.5}
              dot={{ r: 3, fill: colors[i % colors.length], strokeWidth: 0 }}
              activeDot={{ r: 5, strokeWidth: 2, stroke: dark ? '#0f172a' : '#fff' }}
              isAnimationActive={animate} animationDuration={800} />
          ))}
          {showTrendline && yAxis.map((col, i) => (
            <Line key={`trend_${col}`} dataKey={`__trend_${col}`} name={`Trend (${getLabel(col)})`}
              stroke={colors[i % colors.length]} strokeWidth={1.5} strokeDasharray="6 3" dot={false} isAnimationActive={animate} />
          ))}
          {renderBrush()}
        </LineChart>
      </ResponsiveContainer>
    );
  }

  // ── BAR / GROUPED / STACKED ──
  const stackId = chartType === 'stacked' ? 'a' : undefined;
  const bLen = chartData.length;
  const barSize = bLen <= 6 ? 44 : bLen <= 12 ? 32 : bLen <= 24 ? 22 : bLen <= 40 ? 14 : undefined;
  return (
    <ResponsiveContainer width="100%" height={height}>
      <BarChart data={chartData} margin={chartMargin} barCategoryGap="12%">
        <CartesianGrid {...gridProps} />
        <XAxis {...xAxisProps} />
        <YAxis {...yAxisProps} />
        <Tooltip content={<CustomTooltip dark={dark} />} cursor={{ fill: dark ? 'rgba(59,130,246,0.06)' : 'rgba(59,130,246,0.04)' }} />
        <Legend {...legendProps} />
        {yAxis.map((col, i) => (
          hasDiverging && isDivergingField(col)
            ? <Bar key={col} dataKey={col} name={getLabel(col)} stackId={stackId} radius={stackId ? 0 : [6,6,0,0]}
                barSize={barSize} minPointSize={3} isAnimationActive={animate} animationDuration={800}>
                {chartData.map((entry, j) => <Cell key={j} fill={getDivergingColor(entry[col])} />)}
              </Bar>
            : <Bar key={col} dataKey={col} name={getLabel(col)} fill={colors[i % colors.length]}
                stackId={stackId} radius={stackId ? 0 : [6,6,0,0]} barSize={barSize} minPointSize={3}
                isAnimationActive={animate} animationDuration={800} />
        ))}
        {showTrendline && yAxis.map((col, i) => (
          <ReferenceLine key={`ref_${col}`} y={computeStats(chartData.map(r => r[col])).avg}
            stroke={colors[i % colors.length]} strokeDasharray="6 3" strokeWidth={1.5}
            label={{ value: `Avg: ${formatAxisTick(computeStats(chartData.map(r => r[col])).avg)}`, fill: colors[i % colors.length], fontSize: 10, position: 'right' }} />
        ))}
        {renderBrush()}
      </BarChart>
    </ResponsiveContainer>
  );
};



/* ═══════════════════════════════════════════
   STATS RIBBON
   ═══════════════════════════════════════════ */

const StatsRibbon = ({ yColumns, data, dark, palette }) => {
  const colors = COLOR_PALETTES[palette] || COLOR_PALETTES.Ocean;
  return (
    <div className="grid gap-3 mb-5 animate-fadeIn" style={{ gridTemplateColumns: `repeat(${Math.min(yColumns.length, 4)}, 1fr)` }}>
      {yColumns.slice(0, 4).map((col, i) => {
        const stats = computeStats(data.map(r => r[col]));
        const color = colors[i % colors.length];
        return (
          <div key={col}
            className={`rounded-xl border px-4 py-3.5 relative overflow-hidden transition-all duration-300 hover:scale-[1.02] ${
              dark ? 'bg-slate-800/50 border-slate-700/50' : 'bg-white border-slate-200 shadow-sm'
            }`}
          >
            <div className="absolute top-0 left-0 right-0 h-0.5" style={{ backgroundColor: color }} />
            <p className={`text-[10px] font-bold uppercase tracking-wider mb-1 truncate ${dark ? 'text-slate-500' : 'text-slate-400'}`}>{col}</p>
            <p className="text-xl font-extrabold font-mono-data mb-1" style={{ color }}>
              {formatAxisTick(stats.avg)}
            </p>
            <div className="flex items-center gap-2">
              <span className={`text-[10px] font-mono-data ${dark ? 'text-slate-500' : 'text-slate-400'}`}>
                {formatAxisTick(stats.min)} — {formatAxisTick(stats.max)}
              </span>
              <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded-full ${
                stats.trend === 'upward' ? 'bg-emerald-500/15 text-emerald-400'
                : stats.trend === 'downward' ? 'bg-red-500/15 text-red-400'
                : (dark ? 'bg-slate-600/30 text-slate-400' : 'bg-slate-100 text-slate-500')
              }`}>
                {stats.trend === 'upward' ? '↑' : stats.trend === 'downward' ? '↓' : '→'}
                {stats.trendPct !== 0 ? ` ${Math.abs(stats.trendPct).toFixed(1)}%` : ''}
              </span>
            </div>
          </div>
        );
      })}
    </div>
  );
};

/* ═══════════════════════════════════════════
   TOAST / ICONS / REPORT MODAL
   ═══════════════════════════════════════════ */

const Toast = ({ message, type, onClose }) => {
  useEffect(() => { const t = setTimeout(onClose, 3500); return () => clearTimeout(t); }, [onClose]);
  const bg = type === 'error' ? 'bg-red-500/90' : type === 'warning' ? 'bg-amber-500/90' : 'bg-emerald-500/90';
  return (
    <div className={`fixed top-6 right-6 z-[100] ${bg} text-white px-5 py-3 rounded-xl shadow-2xl text-sm font-medium animate-slideUp flex items-center gap-2`}>
      {type === 'error' && <span>⚠️</span>}
      {type === 'warning' && <span>⚡</span>}
      {type === 'success' && <span>✅</span>}
      {message}
    </div>
  );
};

const LogoIcon = () => (
  <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
    <rect x="2" y="2" width="12" height="12" rx="2" fill="#3b82f6" opacity="0.9" />
    <rect x="18" y="2" width="12" height="12" rx="2" fill="#3b82f6" opacity="0.6" />
    <rect x="2" y="18" width="12" height="12" rx="2" fill="#3b82f6" opacity="0.6" />
    <rect x="18" y="18" width="12" height="12" rx="2" fill="#3b82f6" opacity="0.3" />
  </svg>
);
const UploadIcon = () => (
  <svg width="64" height="64" viewBox="0 0 64 64" fill="none" className="mx-auto mb-3 opacity-40">
    <rect x="8" y="20" width="48" height="36" rx="4" stroke="currentColor" strokeWidth="2" fill="none" />
    <path d="M32 16V40M32 16L24 24M32 16L40 24" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    <rect x="16" y="48" width="32" height="4" rx="2" fill="currentColor" opacity="0.2" />
  </svg>
);
const SunIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
    <circle cx="12" cy="12" r="5" /><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
  </svg>
);
const MoonIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
    <path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z" />
  </svg>
);

const ReportModal = ({ charts, data, fieldMeta = {}, onClose, dark }) => {
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  return (
    <div className="fixed inset-0 z-50 overflow-y-auto print:static print:overflow-visible">
      <div className={`min-h-screen ${dark ? 'bg-slate-900' : 'bg-white'} print:bg-white`}>
        <div className="no-print sticky top-0 z-10 flex items-center justify-between px-8 py-4 border-b backdrop-blur-sm"
          style={{ backgroundColor: dark ? 'rgba(15,23,42,0.95)' : 'rgba(255,255,255,0.95)', borderColor: dark ? '#1e293b' : '#e2e8f0' }}>
          <h2 className={`text-lg font-bold ${dark ? 'text-white' : 'text-slate-800'}`}>Report Preview</h2>
          <div className="flex gap-3">
            <button onClick={() => window.print()} className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-500 transition-colors">🖨️ Print / Save as PDF</button>
            <button onClick={onClose} className={`px-5 py-2 rounded-lg text-sm font-semibold border transition-colors ${dark ? 'border-slate-600 text-slate-300 hover:bg-slate-800' : 'border-slate-300 text-slate-600 hover:bg-slate-100'}`}>✕ Close</button>
          </div>
        </div>
        <div className="max-w-4xl mx-auto px-8 py-12 print:px-4 print:py-6">
          <div className="text-center mb-12 print:mb-6">
            <h1 className={`text-3xl font-extrabold mb-2 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Data Insights Report</h1>
            <p className={`text-sm print:text-slate-500 ${dark ? 'text-slate-400' : 'text-slate-500'}`}>{today}</p>
          </div>
          {charts.map((chart, idx) => {
            const processedChartData = processChartData(data, chart.config, fieldMeta);
            const visibleChartData = chart.config.brushStartIndex !== undefined && chart.config.brushEndIndex !== undefined
              ? processedChartData.slice(chart.config.brushStartIndex, chart.config.brushEndIndex + 1)
              : processedChartData;

            const insight = generateInsight(chart.config.yAxis, visibleChartData);
            return (
              <div key={idx} className="report-chart-block mb-16 print:mb-8">
                <h2 className={`text-xl font-bold mb-1 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>
                  {chart.config.yAxis.join(', ')} vs. {chart.config.xAxis}
                </h2>
                <p className={`text-xs mb-5 print:text-slate-500 ${dark ? 'text-slate-500' : 'text-slate-400'}`}>
                  {CHART_TYPES.find(t => t.value === chart.config.chartType)?.label}
                </p>
                <div className={`rounded-xl p-4 mb-5 border print:border-slate-200 ${dark ? 'bg-slate-800/50 border-slate-700' : 'bg-slate-50 border-slate-200'}`}>
                  <ChartRenderer config={chart.config} data={data} fieldMeta={fieldMeta} height={340} dark={dark} animate={false} />
                </div>
                <div className={`rounded-lg p-4 mb-5 border print:bg-slate-50 print:border-slate-200 ${dark ? 'bg-slate-800/30 border-slate-700/50' : 'bg-blue-50/50 border-blue-100'}`}>
                  <p className={`text-sm print:text-slate-700 ${dark ? 'text-slate-300' : 'text-slate-600'}`}>💡 {insight}</p>
                </div>
                <div className="overflow-x-auto">
                  <table className={`w-full text-sm border-collapse rounded-lg overflow-hidden ${dark ? 'text-slate-300' : 'text-slate-600'}`}>
                    <thead>
                      <tr className={`${dark ? 'bg-slate-800' : 'bg-slate-100'} print:bg-slate-100`}>
                        <th className="text-left px-4 py-2.5 font-semibold print:text-slate-700">Metric</th>
                        <th className="text-right px-4 py-2.5 font-semibold font-mono-data print:text-slate-700">Min</th>
                        <th className="text-right px-4 py-2.5 font-semibold font-mono-data print:text-slate-700">Max</th>
                        <th className="text-right px-4 py-2.5 font-semibold font-mono-data print:text-slate-700">Average</th>
                        <th className="text-right px-4 py-2.5 font-semibold print:text-slate-700">Trend</th>
                      </tr>
                    </thead>
                    <tbody>
                      {chart.config.yAxis.map(col => {
                        const stats = computeStats(visibleChartData.map(r => r[col]));
                        return (
                          <tr key={col} className={`border-t ${dark ? 'border-slate-700' : 'border-slate-200'} print:border-slate-200`}>
                            <td className="px-4 py-2.5 font-medium print:text-slate-700">{col}</td>
                            <td className="px-4 py-2.5 text-right font-mono-data print:text-slate-700">{stats.min.toLocaleString()}</td>
                            <td className="px-4 py-2.5 text-right font-mono-data print:text-slate-700">{stats.max.toLocaleString()}</td>
                            <td className="px-4 py-2.5 text-right font-mono-data print:text-slate-700">{stats.avg.toFixed(1)}</td>
                            <td className="px-4 py-2.5 text-right print:text-slate-700">
                              <span className={`inline-flex items-center gap-1 text-xs font-semibold px-2 py-0.5 rounded-full ${
                                stats.trend === 'upward' ? 'bg-emerald-500/20 text-emerald-400 print:text-emerald-700'
                                : stats.trend === 'downward' ? 'bg-red-500/20 text-red-400 print:text-red-700'
                                : (dark ? 'bg-slate-600/30 text-slate-400' : 'bg-slate-200 text-slate-500') + ' print:text-slate-600'
                              }`}>
                                {stats.trend === 'upward' ? '↑' : stats.trend === 'downward' ? '↓' : '→'} {stats.trend}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })}
          <div className={`text-center pt-8 border-t mt-12 print:mt-6 print:border-slate-200 ${dark ? 'border-slate-700 text-slate-500' : 'border-slate-200 text-slate-400'}`}>
            <p className="text-sm print:text-slate-500">Generated by <span className="font-semibold text-blue-400 print:text-blue-600">InsightForge</span></p>
          </div>
        </div>
      </div>
    </div>
  );
};

/* ── Weekly Insights Generator ── */
function generateWeeklyInsights(weeklyReportData, selectedWeeklyMetrics) {
  if (!weeklyReportData) return [];
  const insights = [];
  const { orgRows, selectedIdx } = weeklyReportData;
  
  selectedWeeklyMetrics.forEach(colName => {
    const currentVal = Number(orgRows[selectedIdx][colName]) || 0;
    const prevVal = selectedIdx > 0 
      ? (Number(orgRows[selectedIdx - 1][colName]) || 0) 
      : null;
    const changePct = (prevVal !== null && prevVal > 0)
      ? ((currentVal - prevVal) / prevVal) * 100
      : null;
      
    let text = `${colName} is currently at ${currentVal.toLocaleString()}`;
    if (changePct !== null) {
      text += `, which represents a ${Math.abs(changePct).toFixed(1)}% ${changePct >= 0 ? 'increase' : 'decrease'} compared to the previous week.`;
    } else {
      text += `. This represents the baseline value for the selected period.`;
    }
    
    const lowerCol = colName.toLowerCase();
    if (lowerCol.includes('returning') && lowerCol.includes('mau')) {
      const mauCol = selectedWeeklyMetrics.find(c => c.toLowerCase().includes('current mau') || (c.toLowerCase().includes('mau') && !c.toLowerCase().includes('returning')));
      if (mauCol) {
        const mauVal = Number(orgRows[selectedIdx][mauCol]) || 0;
        if (mauVal > 0) {
          const ratio = (currentVal / mauVal) * 100;
          text += ` Returning users make up ${ratio.toFixed(1)}% of active users, showing a ${ratio > 50 ? 'healthy' : 'moderate'} retention trend.`;
        }
      }
    }
    insights.push(text);
  });

  return insights;
}

/* ── Weekly Report Modal ── */
/* ── Weekly Report Modal ── */
const WeeklyReportModal = ({ orgName, targetDate, weeklyReportData, selectedWeeklyMetrics, weeklyChartType, onClose, dark }) => {
  const [exportOpen, setExportOpen] = React.useState(false);
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const firstRow = weeklyReportData?.historyRows?.[0] || {};
  const regionKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('region')) || 'Focus Region';
  const cohortKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('cohort')) || 'India Cohort';
  const region = firstRow[regionKey] || '—';
  const cohort = firstRow[cohortKey] || '—';

  const colors = [
    { text: 'text-blue-600', hex: '#3b82f6' },
    { text: 'text-violet-600', hex: '#8b5cf6' },
    { text: 'text-emerald-600', hex: '#10b981' },
    { text: 'text-amber-600', hex: '#f59e0b' },
    { text: 'text-rose-600', hex: '#f43f5e' },
    { text: 'text-cyan-600', hex: '#06b6d4' },
    { text: 'text-pink-600', hex: '#ec4899' },
    { text: 'text-indigo-600', hex: '#6366f1' },
  ];

  const metricsList = selectedWeeklyMetrics.map((colName, idx) => {
    const color = colors[idx % colors.length];
    return {
      key: colName,
      label: colName,
      colorClass: color.text,
      colorHex: color.hex
    };
  });

  const insights = generateWeeklyInsights(weeklyReportData, selectedWeeklyMetrics);

  const safeOrgName = orgName.replace(/[^a-zA-Z0-9]/g, '_').replace(/__+/g, '_').replace(/^_+|_+$/g, '');
  const baseFilename = `${safeOrgName}_weekly_report`;

  // ── Excel Export Helper ──
  const exportToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();
      
      // Sheet 1: Executive Summary
      const summaryData = [
        ['WEEKLY ORGANIZATION PERFORMANCE REPORT'],
        [],
        ['Organization Name', orgName],
        ['Week Starting Date', targetDate],
        ['Focus Region', region],
        ['India Cohort', cohort],
        ['Generated Date', today],
        [],
        ['KEY METRICS SUMMARY'],
        ['Metric Name', 'Current Value', 'Previous Value', 'WoW Change %']
      ];
      
      metricsList.forEach(m => {
        const colName = m.key;
        const currentVal = Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx][colName]) || 0;
        const prevVal = weeklyReportData.selectedIdx > 0 
          ? (Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx - 1][colName]) || 0) 
          : null;
        const changePct = (prevVal !== null && prevVal > 0)
          ? ((currentVal - prevVal) / prevVal) * 100
          : null;
        summaryData.push([
          m.label,
          currentVal,
          prevVal !== null ? prevVal : '—',
          changePct !== null ? `${changePct >= 0 ? '+' : ''}${changePct.toFixed(1)}%` : '—'
        ]);
      });
      
      summaryData.push([]);
      summaryData.push(['EXECUTIVE INSIGHTS']);
      insights.forEach(ins => {
        summaryData.push([`• ${ins}`]);
      });
      
      const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
      XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary & Insights');
      
      // Sheet 2: Historical Progression
      const historyData = [];
      const historyHeaders = ['Date'];
      metricsList.forEach(m => {
        historyHeaders.push(m.label);
      });
      historyData.push(['HISTORICAL PROGRESSION DATA']);
      historyData.push(historyHeaders);
      
      weeklyReportData.historyRows.forEach(row => {
        const historyRow = [row[weeklyReportData.dateCol]];
        metricsList.forEach(m => {
          const colName = m.key;
          historyRow.push(row[colName] != null ? Number(row[colName]) : 0);
        });
        historyData.push(historyRow);
      });
      
      const wsHistory = XLSX.utils.aoa_to_sheet(historyData);
      XLSX.utils.book_append_sheet(wb, wsHistory, 'Historical Trend');
      
      // Sheet 3: Rolling Window Details
      const windowData = [
        [`ROLLING ${weeklyReportData.windowSize}-WEEK ACTIVE WINDOW DETAILS`],
        [],
        ['Metric Name', ...weeklyReportData.activeWindow.map(w => `${w.label} (${w.date})`)]
      ];
      
      metricsList.forEach(m => {
        const colName = m.key;
        const rowVals = [m.label];
        weeklyReportData.activeWindow.forEach(week => {
          const val = Number(week.rawRow[colName]) || 0;
          rowVals.push(val);
        });
        windowData.push(rowVals);
      });
      
      const wsWindow = XLSX.utils.aoa_to_sheet(windowData);
      XLSX.utils.book_append_sheet(wb, wsWindow, 'Active Window Details');
      
      XLSX.writeFile(wb, `${baseFilename}.xlsx`);
    } catch (err) {
      console.error('Failed to export Excel:', err);
    }
  };

  // ── PNG Export Helper ──
  const exportToPNG = () => {
    const svgElement = document.querySelector('.weekly-report-chart-container svg');
    if (!svgElement) {
      alert('Trendline chart not found in the DOM.');
      return;
    }
    
    try {
      const clonedSvg = svgElement.cloneNode(true);
      const width = svgElement.clientWidth || 700;
      const height = svgElement.clientHeight || 350;
      clonedSvg.setAttribute('width', width);
      clonedSvg.setAttribute('height', height);
      clonedSvg.style.fontFamily = "'Plus Jakarta Sans', sans-serif";
      clonedSvg.style.backgroundColor = dark ? '#0f172a' : '#ffffff';
      
      const serializer = new XMLSerializer();
      let svgString = serializer.serializeToString(clonedSvg);
      if (!svgString.match(/^<svg[^>]+xmlns="http\:\/\/www\.w3\.org\/2000\/svg"/)) {
        svgString = svgString.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
      }
      
      const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
      const URL = window.URL || window.webkitURL || window;
      const blobURL = URL.createObjectURL(svgBlob);
      
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement('canvas');
        const scale = 2; // high-res
        canvas.width = width * scale;
        canvas.height = height * scale;
        
        const context = canvas.getContext('2d');
        context.scale(scale, scale);
        context.fillStyle = dark ? '#0f172a' : '#ffffff';
        context.fillRect(0, 0, width, height);
        context.drawImage(image, 0, 0, width, height);
        
        const pngUrl = canvas.toDataURL('image/png');
        const downloadLink = document.createElement('a');
        downloadLink.href = pngUrl;
        downloadLink.download = `${baseFilename}.png`;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(blobURL);
      };
      image.onerror = () => {
        const downloadLink = document.createElement('a');
        downloadLink.href = blobURL;
        downloadLink.download = `${baseFilename}.svg`;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
      };
      image.src = blobURL;
    } catch (err) {
      console.error('PNG export failed:', err);
    }
  };

  // ── HTML Export Helper ──
  const exportToHTML = () => {
    const svgElement = document.querySelector('.weekly-report-chart-container svg');
    let svgHtml = '';
    if (svgElement) {
      const clonedSvg = svgElement.cloneNode(true);
      clonedSvg.setAttribute('width', '100%');
      clonedSvg.setAttribute('height', '320');
      clonedSvg.style.backgroundColor = 'transparent';
      
      // Force light mode styling on SVG elements in the exported light-mode HTML
      clonedSvg.querySelectorAll('.recharts-text').forEach(el => {
        el.setAttribute('fill', '#475569');
        el.style.fill = '#475569';
      });
      clonedSvg.querySelectorAll('.recharts-cartesian-grid-horizontal line, .recharts-cartesian-grid-vertical line').forEach(el => {
        el.setAttribute('stroke', '#e2e8f0');
        el.style.stroke = '#e2e8f0';
      });
      
      svgHtml = new XMLSerializer().serializeToString(clonedSvg);
    }
    
    const metricsHtml = metricsList.map(m => {
      const colName = m.key;
      const currentVal = Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx][colName]) || 0;
      const prevVal = weeklyReportData.selectedIdx > 0 
        ? (Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx - 1][colName]) || 0) 
        : null;
      const changePct = (prevVal !== null && prevVal > 0)
        ? ((currentVal - prevVal) / prevVal) * 100
        : null;
      const changeBadge = changePct !== null
        ? `<span style="color: ${changePct >= 0 ? '#059669' : '#dc2626'}; font-weight: 600;">
             ${changePct >= 0 ? '↑' : '↓'} ${Math.abs(changePct).toFixed(1)}%
           </span>`
        : '<span style="color: #94a3b8;">—</span>';
        
      return `
        <tr style="border-top: 1px solid #e2e8f0;">
          <td style="padding: 12px 16px; font-weight: 500; color: #334155;">${m.label}</td>
          <td style="padding: 12px 16px; text-align: right; font-family: monospace; color: #1e293b;">${currentVal.toLocaleString()}</td>
          <td style="padding: 12px 16px; text-align: right; font-family: monospace; color: #475569;">${prevVal !== null ? prevVal.toLocaleString() : '—'}</td>
          <td style="padding: 12px 16px; text-align: right;">${changeBadge}</td>
        </tr>
      `;
    }).join('');
    
    const insightsHtml = insights.map(ins => `
      <li style="margin-bottom: 10px; display: flex; align-items: start; gap: 8px;">
        <span>💡</span>
        <span>${ins}</span>
      </li>
    `).join('');
    
    const rollingWeeksHtml = weeklyReportData.activeWindow.map(week => {
      const isSelectedWeek = week.date === targetDate;
      const weekMetricsHtml = metricsList.map(m => {
        const colName = m.key;
        const val = Number(week.rawRow[colName]) || 0;
        return `
          <div style="display: flex; justify-content: space-between; font-size: 11px; margin-top: 4px;">
            <span style="color: #64748b;">${m.label.replace('Current ', '')}</span>
            <span style="font-weight: 600; color: #334155;">${val.toLocaleString()}</span>
          </div>
        `;
      }).join('');
      
      return `
        <div style="flex: 1; min-width: 160px; border: 1px solid ${isSelectedWeek ? '#8b5cf6' : '#e2e8f0'}; background-color: ${isSelectedWeek ? '#f5f3ff' : '#ffffff'}; padding: 12px; border-radius: 12px;">
          <div style="font-size: 10px; font-weight: bold; color: ${isSelectedWeek ? '#7c3aed' : '#64748b'}; text-transform: uppercase;">
            ${week.label} ${isSelectedWeek ? '(Target)' : ''}
          </div>
          <div style="font-family: monospace; font-size: 11px; font-weight: 600; margin-bottom: 8px; color: #475569;">
            ${new Date(week.date).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
          </div>
          <div>${weekMetricsHtml}</div>
        </div>
      `;
    }).join('');
    
    const fullHtml = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Weekly Org Performance Report - ${orgName}</title>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
      background-color: #f8fafc;
      margin: 0;
      padding: 24px;
      color: #0f172a;
    }
    .container {
      max-width: 900px;
      margin: 0 auto;
      background: white;
      padding: 40px;
      border-radius: 16px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
      border: 1px solid #e2e8f0;
    }
    .header {
      text-align: center;
      margin-bottom: 32px;
      border-bottom: 1px solid #e2e8f0;
      padding-bottom: 24px;
    }
    .title {
      font-size: 28px;
      font-weight: 800;
      color: #1e293b;
      margin: 0 0 8px 0;
    }
    .subtitle {
      font-size: 16px;
      color: #64748b;
      margin: 0 0 8px 0;
      font-weight: 500;
    }
    .meta-tag {
      display: inline-block;
      font-size: 10px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      padding: 4px 10px;
      border-radius: 8px;
      background-color: #f1f5f9;
      color: #475569;
      margin: 0 4px;
      border: 1px solid #e2e8f0;
    }
    .meta-tag-purple {
      background-color: #f5f3ff;
      color: #6d28d9;
      border: 1px solid #ddd6fe;
    }
    .section-title {
      font-size: 18px;
      font-weight: 700;
      color: #1e293b;
      margin: 24px 0 12px 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 24px;
      border-radius: 8px;
      overflow: hidden;
      border: 1px solid #e2e8f0;
    }
    th {
      background-color: #f1f5f9;
      color: #475569;
      text-align: left;
      padding: 10px 16px;
      font-size: 12px;
      font-weight: 600;
    }
    .chart-container {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      padding: 20px;
      border-radius: 12px;
      margin-bottom: 24px;
      height: 320px;
    }
    .insights-box {
      background-color: #eff6ff;
      border: 1px solid #dbeafe;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 24px;
    }
    .insights-list {
      list-style: none;
      padding: 0;
      margin: 0;
      font-size: 13px;
      color: #334155;
      line-height: 1.6;
    }
    .rolling-grid {
      display: flex;
      gap: 16px;
      flex-wrap: wrap;
      margin-bottom: 24px;
    }
    .transition-box {
      background-color: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      padding: 16px;
      font-size: 12px;
      color: #475569;
    }
    .footer {
      text-align: center;
      margin-top: 40px;
      border-top: 1px solid #e2e8f0;
      padding-top: 20px;
      font-size: 11px;
      color: #94a3b8;
    }
    @media print {
      body { background: white; padding: 0; }
      .container { border: none; box-shadow: none; padding: 0; }
      .chart-container { page-break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 class="title">Weekly Organization Performance Report</h1>
      <p class="subtitle">${orgName}</p>
      <div style="margin-bottom: 12px;">
        <span class="meta-tag">Region: ${region}</span>
        <span class="meta-tag meta-tag-purple">Cohort: ${cohort}</span>
      </div>
      <p style="font-size: 11px; color: #64748b; margin: 0;">
        Week starting: <strong>${targetDate}</strong> · Generated on ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}
      </p>
    </div>
    
    <div class="section-title">Key Metrics Summary</div>
    <table>
      <thead>
        <tr>
          <th style="text-align: left;">Metric Name</th>
          <th style="text-align: right;">Current Value</th>
          <th style="text-align: right;">Previous Value</th>
          <th style="text-align: right;">WoW Change</th>
        </tr>
      </thead>
      <tbody>
        ${metricsHtml}
      </tbody>
    </table>
    
    <div class="section-title">Historical Trend Progression (${weeklyChartType.toUpperCase()})</div>
    <div class="chart-container">
      ${svgHtml}
    </div>
    
    <div class="section-title">Weekly Executive Insights</div>
    <div class="insights-box">
      <ul class="insights-list">
        ${insightsHtml}
      </ul>
    </div>
    
    <div class="section-title">Rolling ${weeklyReportData.windowSize}-Week Active Window Details</div>
    <div class="rolling-grid">
      ${rollingWeeksHtml}
    </div>
    
    <div class="section-title">Rolling Window Transition</div>
    <div class="transition-box">
      <p style="margin: 0 0 8px 0;">To compute the rolling ${weeklyReportData.windowSize === 12 ? 'Quarter' : 'Monthly Active Users (MAU)'} as of <strong>${targetDate}</strong>:</p>
      <ul style="margin: 0; padding-left: 20px;">
        ${weeklyReportData.droppedWeekDate 
          ? `<li style="margin-bottom: 4px;">Oldest week <span style="color: #ef4444; font-weight: 600;">${weeklyReportData.droppedWeekDate}</span> was dropped from the rolling window.</li>`
          : '<li style="margin-bottom: 4px;">Initial window setup phase: no week dropped yet.</li>'}
        <li>New week <span style="color: #10b981; font-weight: 600;">${targetDate}</span> was added.</li>
      </ul>
    </div>
    
    <div class="footer">
      Generated by <strong>InsightForge Dashboard</strong>
    </div>
  </div>
</body>
</html>`;

    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const htmlUrl = URL.createObjectURL(blob);
    const downloadLink = document.createElement('a');
    downloadLink.href = htmlUrl;
    downloadLink.download = `${baseFilename}.html`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(htmlUrl);
  };

  // ── TXT Export Helper ──
  const exportToTXT = () => {
    let txt = `========================================================\n`;
    txt += `WEEKLY ORGANIZATION PERFORMANCE REPORT\n`;
    txt += `========================================================\n\n`;
    txt += `Organization Name : ${orgName}\n`;
    txt += `Target Date       : ${targetDate}\n`;
    txt += `Focus Region      : ${region}\n`;
    txt += `India Cohort      : ${cohort}\n`;
    txt += `Generated On      : ${today}\n\n`;
    
    txt += `--------------------------------------------------------\n`;
    txt += `KEY METRICS SUMMARY\n`;
    txt += `--------------------------------------------------------\n`;
    txt += `${'Metric Name'.padEnd(25)} ${'Current Val'.padEnd(15)} ${'Previous Val'.padEnd(15)} ${'WoW Change'.padEnd(15)}\n`;
    
    metricsList.forEach(m => {
      const colName = m.key;
      const currentVal = Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx][colName]) || 0;
      const prevVal = weeklyReportData.selectedIdx > 0 
        ? (Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx - 1][colName]) || 0) 
        : null;
      const changePct = (prevVal !== null && prevVal > 0)
        ? ((currentVal - prevVal) / prevVal) * 100
        : null;
      
      const currentStr = currentVal.toLocaleString();
      const prevStr = prevVal !== null ? prevVal.toLocaleString() : '—';
      const changeStr = changePct !== null ? `${changePct >= 0 ? '↑' : '↓'} ${Math.abs(changePct).toFixed(1)}%` : '—';
      
      txt += `${m.label.padEnd(25)} ${currentStr.padEnd(15)} ${prevStr.padEnd(15)} ${changeStr.padEnd(15)}\n`;
    });
    
    txt += `\n--------------------------------------------------------\n`;
    txt += `WEEKLY EXECUTIVE INSIGHTS\n`;
    txt += `--------------------------------------------------------\n`;
    insights.forEach(ins => {
      txt += `• ${ins}\n`;
    });
    
    txt += `\n--------------------------------------------------------\n`;
    txt += `ROLLING ${weeklyReportData.windowSize}-WEEK ACTIVE WINDOW DETAILS\n`;
    txt += `--------------------------------------------------------\n`;
    weeklyReportData.activeWindow.forEach(week => {
      const isSelected = week.date === targetDate;
      txt += `${week.label} ${isSelected ? '(Target)' : ''} [${week.date}]:\n`;
      metricsList.forEach(m => {
        const colName = m.key;
        const val = Number(week.rawRow[colName]) || 0;
        txt += `  - ${m.label.replace('Current ', '')}: ${val.toLocaleString()}\n`;
      });
    });
    
    txt += `\n--------------------------------------------------------\n`;
    txt += `ROLLING WINDOW TRANSITION DETAILS\n`;
    txt += `--------------------------------------------------------\n`;
    if (weeklyReportData.droppedWeekDate) {
      txt += `- Oldest week [${weeklyReportData.droppedWeekDate}] was dropped.\n`;
    } else {
      txt += `- Initial window setup: no week dropped yet.\n`;
    }
    txt += `- New week [${targetDate}] was added.\n\n`;
    txt += `Report generated by InsightForge\n`;
    
    const blob = new Blob([txt], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement('a');
    downloadLink.href = url;
    downloadLink.download = `${baseFilename}.txt`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto print:static print:overflow-visible">
      <style dangerouslySetInnerHTML={{ __html: `
        @media print {
          body, html, #root {
            background-color: #ffffff !important;
            color: #0f172a !important;
          }
          /* Force light mode styles on all elements during print */
          .print\\:bg-white { background-color: #ffffff !important; }
          .print\\:text-black { color: #000000 !important; }
          .print\\:text-slate-700 { color: #334155 !important; }
          .print\\:border-slate-200 { border-color: #e2e8f0 !important; }
          
          /* Ensure text colors are dark for print readability */
          span, p, h1, h2, h3, h4, td, th, li, div {
            color: #0f172a !important;
          }
          
          .text-slate-400, .text-slate-500, .text-slate-600, .text-slate-300, .text-slate-200 {
            color: #334155 !important;
          }
          
          /* Keep specific metric badge colors if needed, but make them print-friendly */
          .text-blue-600, .text-blue-500 { color: #1d4ed8 !important; }
          .text-violet-600, .text-violet-500 { color: #6d28d9 !important; }
          .text-emerald-600, .text-emerald-500 { color: #047857 !important; }
          .text-amber-600, .text-amber-500 { color: #b45309 !important; }
          
          /* Ensure the chart SVG text is black/slate for print */
          .recharts-text {
            fill: #475569 !important;
          }
          .recharts-cartesian-grid-horizontal line, 
          .recharts-cartesian-grid-vertical line {
            stroke: #e2e8f0 !important;
          }
        }
      `}} />
      <div className={`min-h-screen ${dark ? 'bg-slate-900' : 'bg-white'} print:bg-white`}>
        <div className="no-print sticky top-0 z-10 flex items-center justify-between px-8 py-4 border-b backdrop-blur-sm"
          style={{ backgroundColor: dark ? 'rgba(15,23,42,0.95)' : 'rgba(255,255,255,0.95)', borderColor: dark ? '#1e293b' : '#e2e8f0' }}>
          <h2 className={`text-lg font-bold ${dark ? 'text-white' : 'text-slate-800'}`}>Weekly Report Preview</h2>
          <div className="flex gap-3 relative">
            <button 
              onClick={() => {
                const originalTitle = document.title;
                document.title = baseFilename;
                window.print();
                document.title = originalTitle;
              }} 
              className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-semibold hover:bg-blue-500 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <span>🖨️</span> Print / Save as PDF
            </button>
            
            <div className="relative">
              <button 
                onClick={() => setExportOpen(!exportOpen)} 
                className={`px-5 py-2 rounded-lg text-sm font-semibold border transition-colors flex items-center gap-1.5 shadow-sm ${
                  dark 
                    ? 'border-slate-700 bg-slate-800 text-slate-200 hover:bg-slate-700' 
                    : 'border-slate-300 bg-white text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span>📥</span> Export As <span className="text-[9px] opacity-75">▼</span>
              </button>
              
              {exportOpen && (
                <>
                  <div className="fixed inset-0 z-30" onClick={() => setExportOpen(false)} />
                  <div className={`absolute right-0 mt-2 w-56 rounded-xl border shadow-xl z-40 py-1.5 animate-fadeIn ${
                    dark 
                      ? 'bg-slate-800 border-slate-700 text-slate-200' 
                      : 'bg-white border-slate-200 text-slate-700'
                  }`}>
                    <button 
                      onClick={() => { exportToExcel(); setExportOpen(false); }} 
                      className={`w-full text-left px-4 py-2.5 text-xs font-semibold hover:bg-violet-600 hover:text-white transition-colors flex items-center gap-2`}
                    >
                      <span>📊</span> Excel Spreadsheet (.xlsx)
                    </button>
                    <button 
                      onClick={() => { exportToHTML(); setExportOpen(false); }} 
                      className={`w-full text-left px-4 py-2.5 text-xs font-semibold hover:bg-violet-600 hover:text-white transition-colors flex items-center gap-2`}
                    >
                      <span>🌐</span> Standalone HTML Report (.html)
                    </button>
                    <button 
                      onClick={() => { exportToPNG(); setExportOpen(false); }} 
                      className={`w-full text-left px-4 py-2.5 text-xs font-semibold hover:bg-violet-600 hover:text-white transition-colors flex items-center gap-2`}
                    >
                      <span>📷</span> Trendline Chart Image (.png)
                    </button>
                    <button 
                      onClick={() => { exportToTXT(); setExportOpen(false); }} 
                      className={`w-full text-left px-4 py-2.5 text-xs font-semibold hover:bg-violet-600 hover:text-white transition-colors flex items-center gap-2`}
                    >
                      <span>📄</span> Executive Insights (.txt)
                    </button>
                  </div>
                </>
              )}
            </div>
            
            <button onClick={onClose} className={`px-5 py-2 rounded-lg text-sm font-semibold border transition-colors ${dark ? 'border-slate-600 text-slate-300 hover:bg-slate-800' : 'border-slate-300 text-slate-600 hover:bg-slate-100'}`}>✕ Close</button>
          </div>
        </div>
        
        <div className="max-w-4xl mx-auto px-8 py-12 print:px-4 print:py-6">
          {/* Header block */}
          <div className="text-center mb-10 print:mb-6">
            <h1 className={`text-3xl font-extrabold mb-2 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Weekly Organization Performance Report</h1>
            <p className={`text-sm print:text-slate-500 ${dark ? 'text-slate-400' : 'text-slate-500'}`}>{orgName}</p>
            <p className={`text-xs mt-1 print:text-slate-400 ${dark ? 'text-slate-500' : 'text-slate-400'}`}>Week starting: <span className="font-semibold text-violet-500">{targetDate}</span> · Generated on {today}</p>
          </div>

          {/* Metadata */}
          <div className={`rounded-xl p-4 mb-6 border print:border-slate-200 ${dark ? 'bg-slate-800/40 border-slate-700/50' : 'bg-slate-50 border-slate-200'}`}>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <div>
                <span className={`block text-xs font-semibold ${dark ? 'text-slate-500' : 'text-slate-400'}`}>Focus Region</span>
                <span className={`font-bold ${dark ? 'text-slate-200' : 'text-slate-800'}`}>{region}</span>
              </div>
              <div>
                <span className={`block text-xs font-semibold ${dark ? 'text-slate-500' : 'text-slate-400'}`}>India Cohort</span>
                <span className={`font-bold ${dark ? 'text-slate-200' : 'text-slate-800'}`}>{cohort}</span>
              </div>
            </div>
          </div>

          {/* Key Metrics Table */}
          <div className="mb-8">
            <h2 className={`text-lg font-bold mb-3 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Key Metrics Summary</h2>
            <div className="overflow-x-auto rounded-lg border border-slate-200">
              <table className={`w-full text-sm border-collapse ${dark ? 'text-slate-300' : 'text-slate-600'}`}>
                <thead>
                  <tr className={`${dark ? 'bg-slate-800' : 'bg-slate-100'} print:bg-slate-100`}>
                    <th className="text-left px-4 py-2.5 font-semibold print:text-slate-700">Metric Name</th>
                    <th className="text-right px-4 py-2.5 font-semibold print:text-slate-700 font-mono-data">Current Value</th>
                    <th className="text-right px-4 py-2.5 font-semibold print:text-slate-700 font-mono-data">Previous Value</th>
                    <th className="text-right px-4 py-2.5 font-semibold print:text-slate-700">WoW Change</th>
                  </tr>
                </thead>
                <tbody>
                  {metricsList.map(m => {
                    const colName = m.key;
                    const currentVal = Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx][colName]) || 0;
                    const prevVal = weeklyReportData.selectedIdx > 0 
                      ? (Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx - 1][colName]) || 0) 
                      : null;
                    const changePct = (prevVal !== null && prevVal > 0)
                      ? ((currentVal - prevVal) / prevVal) * 100
                      : null;

                    return (
                      <tr key={m.key} className={`border-t ${dark ? 'border-slate-700' : 'border-slate-200'} print:border-slate-200`}>
                        <td className="px-4 py-2.5 font-medium print:text-slate-700">{m.label}</td>
                        <td className="px-4 py-2.5 text-right font-mono-data print:text-slate-700">{currentVal.toLocaleString()}</td>
                        <td className="px-4 py-2.5 text-right font-mono-data print:text-slate-700">{prevVal !== null ? prevVal.toLocaleString() : '—'}</td>
                        <td className="px-4 py-2.5 text-right font-semibold">
                          {changePct !== null ? (
                            <span className={changePct >= 0 ? 'text-emerald-600' : 'text-red-500'}>
                              {changePct >= 0 ? '↑' : '↓'} {Math.abs(changePct).toFixed(1)}%
                            </span>
                          ) : (
                            <span className={dark ? 'text-slate-500' : 'text-slate-400'}>—</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Chart block */}
          <div className="mb-8 page-break-before">
            <h2 className={`text-lg font-bold mb-3 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Historical Trend Progression ({weeklyChartType.toUpperCase()})</h2>
            <div className={`rounded-xl p-5 border print:border-slate-200 ${dark ? 'bg-slate-800/40 border-slate-700' : 'bg-slate-50 border-slate-200'}`}>
              
              {/* Screen-only Responsive Chart */}
              <div className="print:hidden weekly-report-chart-container" style={{ width: '100%', height: 320 }}>
                <ResponsiveContainer>
                  <ComposedChart data={weeklyReportData.historyRows}>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark ? '#334155' : '#e2e8f0'} vertical={false} />
                    <XAxis 
                      dataKey={weeklyReportData.dateCol} 
                      stroke={dark ? '#64748b' : '#94a3b8'} 
                      fontSize={10} 
                      tickLine={false} 
                      axisLine={false}
                      tickFormatter={(str) => {
                        const d = new Date(str);
                        return isNaN(d) ? str : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                      }}
                    />
                    <YAxis 
                      stroke={dark ? '#64748b' : '#94a3b8'} 
                      fontSize={10} 
                      tickLine={false} 
                      axisLine={false}
                      tickFormatter={(val) => val.toLocaleString()}
                    />
                    <Tooltip content={<CustomTooltip dark={dark} />} />
                    <Legend iconType="circle" wrapperStyle={{ fontSize: 11, paddingTop: 10 }} />
                    
                    <ReferenceLine 
                      x={targetDate} 
                      stroke={dark ? '#a855f7' : '#8b5cf6'} 
                      strokeWidth={2}
                      strokeDasharray="4 4"
                      label={{ value: 'Target Week', fill: dark ? '#c084fc' : '#6b21a8', fontSize: 10, position: 'top' }} 
                    />
                    
                    {metricsList.map(m => {
                      const colName = m.key;
                      const color = m.colorHex;
                      if (weeklyChartType === 'bar') {
                        return (
                          <Bar 
                            key={m.key} 
                            dataKey={colName} 
                            name={m.label} 
                            fill={color} 
                            radius={[4, 4, 0, 0]}
                          />
                        );
                      } else if (weeklyChartType === 'area') {
                        return (
                          <Area 
                            key={m.key} 
                            type="monotone"
                            dataKey={colName} 
                            name={m.label} 
                            stroke={color} 
                            fill={color} 
                            fillOpacity={0.15}
                            strokeWidth={2}
                          />
                        );
                      } else {
                        return (
                          <Line 
                            key={m.key} 
                            type="monotone" 
                            dataKey={colName} 
                            name={m.label} 
                            stroke={color} 
                            strokeWidth={3} 
                            dot={{ r: 4, strokeWidth: 1 }}
                            activeDot={{ r: 7, strokeWidth: 0 }}
                          />
                        );
                      }
                    })}
                  </ComposedChart>
                </ResponsiveContainer>
              </div>

              {/* Print-only Static Chart (Prevents Recharts print container collapse) */}
              <div className="hidden print:block weekly-report-chart-container-print" style={{ width: 680, height: 320 }}>
                <ComposedChart width={680} height={320} data={weeklyReportData.historyRows} margin={{ top: 20, right: 30, bottom: 20, left: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                  <XAxis 
                    dataKey={weeklyReportData.dateCol} 
                    stroke="#94a3b8" 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false}
                    tickFormatter={(str) => {
                      const d = new Date(str);
                      return isNaN(d) ? str : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                    }}
                  />
                  <YAxis 
                    stroke="#94a3b8" 
                    fontSize={10} 
                    tickLine={false} 
                    axisLine={false}
                    tickFormatter={(val) => val.toLocaleString()}
                  />
                  <Legend iconType="circle" wrapperStyle={{ fontSize: 11, paddingTop: 10 }} />
                  
                  <ReferenceLine 
                    x={targetDate} 
                    stroke="#8b5cf6" 
                    strokeWidth={2}
                    strokeDasharray="4 4"
                    label={{ value: 'Target Week', fill: '#6b21a8', fontSize: 10, position: 'top' }} 
                  />
                  
                  {metricsList.map(m => {
                    const colName = m.key;
                    const color = m.colorHex;
                    if (weeklyChartType === 'bar') {
                      return (
                        <Bar 
                          key={m.key} 
                          dataKey={colName} 
                          name={m.label} 
                          fill={color} 
                          radius={[4, 4, 0, 0]}
                        />
                      );
                    } else if (weeklyChartType === 'area') {
                      return (
                        <Area 
                          key={m.key} 
                          type="monotone"
                          dataKey={colName} 
                          name={m.label} 
                          stroke={color} 
                          fill={color} 
                          fillOpacity={0.15}
                          strokeWidth={2}
                        />
                      );
                    } else {
                      return (
                        <Line 
                          key={m.key} 
                          type="monotone" 
                          dataKey={colName} 
                          name={m.label} 
                          stroke={color} 
                          strokeWidth={3} 
                          dot={{ r: 4, strokeWidth: 1 }}
                        />
                      );
                    }
                  })}
                </ComposedChart>
              </div>

            </div>
          </div>

          {/* Automated insights block */}
          <div className="mb-8 print:border-slate-200">
            <h2 className={`text-lg font-bold mb-3 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Weekly Executive Insights</h2>
            <div className={`rounded-lg p-5 border leading-relaxed ${dark ? 'bg-slate-800/30 border-slate-700/50 text-slate-300' : 'bg-blue-50/50 border-blue-100 text-slate-700'}`}>
              <ul className="list-disc pl-5 space-y-2.5 text-sm">
                {insights.map((ins, i) => (
                  <li key={i}>💡 {ins}</li>
                ))}
              </ul>
            </div>
          </div>

          {/* Rolling window detail */}
          <div className="mb-8 page-break-before">
            <h2 className={`text-lg font-bold mb-3 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Rolling {weeklyReportData.windowSize}-Week Active Window Details</h2>
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-4">
              {weeklyReportData.activeWindow.map((week) => {
                const isSelectedWeek = week.date === targetDate;
                return (
                  <div key={week.label} className={`rounded-xl border p-4 ${isSelectedWeek ? 'border-violet-500 bg-violet-500/5' : 'border-slate-200'}`}>
                    <p className="text-xs font-bold text-slate-500 mb-1">{week.label} {isSelectedWeek && '(Target)'}</p>
                    <p className="text-xs font-mono font-semibold mb-2">{new Date(week.date).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}</p>
                    <div className="space-y-1">
                      {metricsList.map(m => {
                        const colName = m.key;
                        const val = Number(week.rawRow[colName]) || 0;
                        return (
                          <div key={m.key} className="flex justify-between text-xs">
                            <span className="text-slate-400">{m.label.replace('Current ', '')}</span>
                            <span className="font-semibold text-slate-700">{val.toLocaleString()}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Transition details */}
          <div className="mb-8">
            <h2 className={`text-lg font-bold mb-3 print:text-black ${dark ? 'text-white' : 'text-slate-800'}`}>Rolling Window Transition</h2>
            <div className={`rounded-xl p-4 border text-sm ${dark ? 'bg-slate-800/10 border-slate-700' : 'bg-slate-50 border-slate-200'}`}>
              <p className={dark ? 'text-slate-400' : 'text-slate-600'}>
                To compute the rolling {weeklyReportData.windowSize === 12 ? 'Quarter' : 'Monthly Active Users (MAU)'} as of <span className="font-semibold font-mono">{targetDate}</span>:
              </p>
              <ul className={`list-disc pl-5 mt-2 space-y-1 text-xs ${dark ? 'text-slate-500' : 'text-slate-500'}`}>
                {weeklyReportData.droppedWeekDate ? (
                  <li>Oldest week <span className="text-red-500 font-semibold font-mono">{weeklyReportData.droppedWeekDate}</span> was dropped.</li>
                ) : (
                  <li>Initial database setup phase: no week dropped yet.</li>
                )}
                <li>New week <span className="text-emerald-500 font-semibold font-mono">{targetDate}</span> was added.</li>
              </ul>
            </div>
          </div>

          <div className={`text-center pt-8 border-t mt-12 print:mt-6 print:border-slate-200 ${dark ? 'border-slate-700 text-slate-500' : 'border-slate-200 text-slate-400'}`}>
            <p className="text-sm print:text-slate-500">Generated by <span className="font-semibold text-blue-400 print:text-blue-600">InsightForge</span></p>
          </div>
        </div>
      </div>
    </div>
  );
};

/* ═══════════════════════════════════════════
   MAIN APP COMPONENT
   ═══════════════════════════════════════════ */

export default function App() {
  const [dark, setDark] = useState(false);

  // CSV state
  const [csvData, setCsvData] = useState(null);
  const [columns, setColumns] = useState([]);
  const [fieldMeta, setFieldMeta] = useState({});
  const [fileName, setFileName] = useState('');
  const [rowCount, setRowCount] = useState(0);
  const [previewRows, setPreviewRows] = useState([]);
  const [loading, setLoading] = useState(false);

  // Weekly Report state
  const [viewMode, setViewMode] = useState('weekly');
  const [selectedOrg, setSelectedOrg] = useState('');
  const [selectedWeeklyStartDate, setSelectedWeeklyStartDate] = useState('');
  const [selectedWeeklyEndDate, setSelectedWeeklyEndDate] = useState('');
  const [selectedWeeklyMetrics, setSelectedWeeklyMetrics] = useState([]);
  const [weeklyHistoryLimit, setWeeklyHistoryLimit] = useState('mau'); // 'mau' or 'quarter'
  const [weeklyChartType, setWeeklyChartType] = useState('line');

  // Chart config
  const [xAxis, setXAxis] = useState('');
  const [yAxis, setYAxis] = useState([]);
  const [chartType, setChartType] = useState('bar');
  const [palette, setPalette] = useState('Ocean');
  const [showTrendline, setShowTrendline] = useState(false);
  const [legendLabels, setLegendLabels] = useState({});
  const [showPreview, setShowPreview] = useState(false);
  const [filterZeros, setFilterZeros] = useState(true);
  const [topN, setTopN] = useState('25');
  const [logScale, setLogScale] = useState(false);
  const [aggregation, setAggregation] = useState('sum');
  const [bubbleField, setBubbleField] = useState('');
  const [brushStartIndex, setBrushStartIndex] = useState(0);
  const [brushEndIndex, setBrushEndIndex] = useState(24);

  // Report queue
  const [savedCharts, setSavedCharts] = useState([]);
  const [showReport, setShowReport] = useState(false);
  const [showWeeklyReport, setShowWeeklyReport] = useState(false);

  const [toast, setToast] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const fileInputRef = useRef(null);

  // ── Processed data calculation ──
  const processedData = useMemo(() => {
    return processChartData(csvData, { xAxis, yAxis, chartType, filterZeros, topN, aggregation, bubbleField }, fieldMeta);
  }, [csvData, xAxis, yAxis, chartType, filterZeros, topN, aggregation, bubbleField, fieldMeta]);

  // ── Auto reset brush indices when data changes ──
  useEffect(() => {
    setBrushStartIndex(0);
    setBrushEndIndex(Math.max(0, Math.min(processedData.length - 1, 24)));
  }, [processedData]);

  // ── Sliced visible data for stats/insights ──
  const visibleData = useMemo(() => {
    if (processedData.length === 0) return [];
    const hasBrush = !PIE_TYPES.includes(chartType) && chartType !== 'horizontal' && chartType !== 'scatter' && chartType !== 'bubble' && processedData.length > 25;
    if (hasBrush) {
      return processedData.slice(brushStartIndex, brushEndIndex + 1);
    }
    return processedData;
  }, [processedData, brushStartIndex, brushEndIndex, chartType]);

  // Extract unique organizations from loaded CSV
  const uniqueOrgs = useMemo(() => {
    if (!csvData || csvData.length === 0) return [];
    const orgCol = getOrgNameColumn(csvData[0]);
    if (!orgCol) return [];
    return Array.from(new Set(csvData.map(r => String(r[orgCol])).filter(Boolean))).sort();
  }, [csvData]);

  // Extract unique dates specifically for the selected organization
  const uniqueOrgDates = useMemo(() => {
    if (!csvData || csvData.length === 0 || !selectedOrg) return [];
    const orgCol = getOrgNameColumn(csvData[0]);
    const dateCol = getAsOfDateColumn(csvData[0]);
    if (!orgCol || !dateCol) return [];
    return Array.from(new Set(
      csvData
        .filter(r => String(r[orgCol]) === selectedOrg)
        .map(r => String(r[dateCol]))
        .filter(Boolean)
    )).sort((a, b) => new Date(a) - new Date(b));
  }, [csvData, selectedOrg]);

  // Extract unique dates globally from loaded CSV
  const uniqueDates = useMemo(() => {
    if (!csvData || csvData.length === 0) return [];
    const dateCol = getAsOfDateColumn(csvData[0]);
    if (!dateCol) return [];
    return Array.from(new Set(csvData.map(r => String(r[dateCol])).filter(Boolean))).sort((a, b) => new Date(a) - new Date(b));
  }, [csvData]);

  // Auto-select first organization on file load
  useEffect(() => {
    if (uniqueOrgs.length > 0 && (!selectedOrg || !uniqueOrgs.includes(selectedOrg))) {
      setSelectedOrg(uniqueOrgs[0]);
    }
  }, [uniqueOrgs, selectedOrg]);

  // Auto-select date range on organization load
  useEffect(() => {
    if (uniqueOrgDates.length > 0) {
      setSelectedWeeklyStartDate(uniqueOrgDates[0]);
      if (weeklyHistoryLimit === 'mau') {
        const targetEndIdx = Math.min(uniqueOrgDates.length - 1, 3);
        setSelectedWeeklyEndDate(uniqueOrgDates[targetEndIdx]);
      } else {
        setSelectedWeeklyEndDate(uniqueOrgDates[uniqueOrgDates.length - 1]);
      }
    } else {
      setSelectedWeeklyStartDate('');
      setSelectedWeeklyEndDate('');
    }
  }, [uniqueOrgDates, weeklyHistoryLimit]);



  // Weekly report data processing
  const weeklyReportData = useMemo(() => {
    if (!csvData || csvData.length === 0 || !selectedOrg || !selectedWeeklyStartDate || !selectedWeeklyEndDate) return null;
    
    const orgCol = getOrgNameColumn(csvData[0]);
    const dateCol = getAsOfDateColumn(csvData[0]);
    if (!orgCol || !dateCol) return null;

    // Filter and sort chronological data for selected organization
    const orgRows = csvData
      .filter(r => String(r[orgCol]) === selectedOrg)
      .sort((a, b) => new Date(a[dateCol]) - new Date(b[dateCol]));

    if (orgRows.length === 0) return null;

    // Find the indices of the selected start and end dates
    const startIdx = orgRows.findIndex(r => String(r[dateCol]) === selectedWeeklyStartDate);
    const endIdx = orgRows.findIndex(r => String(r[dateCol]) === selectedWeeklyEndDate);
    if (startIdx === -1 || endIdx === -1) return null;

    const chronologicalSlice = orgRows.slice(Math.min(startIdx, endIdx), Math.max(startIdx, endIdx) + 1);
    const historyRows = chronologicalSlice;

    // Rolling window details (MAU = 4 weeks, Quarter = 12 weeks)
    const targetIdx = Math.max(startIdx, endIdx);
    const windowSize = weeklyHistoryLimit === 'quarter' ? 12 : 4;
    const windowStartIdx = Math.max(0, targetIdx - windowSize + 1);
    const windowRows = orgRows.slice(windowStartIdx, targetIdx + 1);
    
    // Map active window rows with W1, W2, etc. labels
    const activeWindow = windowRows.map((row, i, arr) => {
      const weekNum = windowSize - (arr.length - 1 - i);
      return {
        label: `W${weekNum}`,
        date: row[dateCol],
        rawRow: row
      };
    });

    // Dropped vs Added transition details
    const droppedWeekRow = targetIdx >= windowSize ? orgRows[targetIdx - windowSize] : null;
    const droppedWeekDate = droppedWeekRow ? droppedWeekRow[dateCol] : null;
    const addedWeekDate = selectedWeeklyEndDate;

    return {
      orgCol,
      dateCol,
      orgRows,
      selectedIdx: targetIdx,
      historyRows,
      activeWindow,
      droppedWeekRow,
      droppedWeekDate,
      addedWeekDate,
      windowSize
    };
  }, [csvData, selectedOrg, selectedWeeklyStartDate, selectedWeeklyEndDate, weeklyHistoryLimit]);

  // ── Smart recommendations ──
  const recommendations = useMemo(() => {
    if (!csvData || !xAxis || yAxis.length === 0) return [];
    return getChartRecommendations(xAxis, yAxis, fieldMeta, csvData, columns);
  }, [csvData, xAxis, yAxis, fieldMeta, columns]);

  // ── Load parsed data into state ──
  const loadData = useCallback((data, name) => {
    if (!data || data.length === 0) {
      setToast({ message: 'File appears empty or invalid', type: 'error' });
      setLoading(false);
      return;
    }
    const cols = Object.keys(data[0]);
    const meta = getFieldMetadata(cols, data);
    setCsvData(data);
    setColumns(cols);
    setFieldMeta(meta);
    setFileName(name);
    setRowCount(data.length);
    setPreviewRows(data.slice(0, 3));
    setXAxis(cols[0] || '');
    setYAxis([]);
    setShowPreview(false);
    setViewMode('weekly');
    setLoading(false);
    setToast({ message: `Loaded ${data.length} rows from ${name}`, type: 'success' });
  }, []);

  const handleParseCSV = useCallback((text, name) => {
    setLoading(true);
    setTimeout(() => {
      Papa.parse(text, {
        header: true, skipEmptyLines: true, dynamicTyping: true,
        complete: (result) => loadData(result.data, name),
        error: () => { setToast({ message: 'Failed to parse CSV', type: 'error' }); setLoading(false); },
      });
    }, 300);
  }, [loadData]);

  const handleParseXLSX = useCallback((buffer, name) => {
    setLoading(true);
    setTimeout(() => {
      try {
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        loadData(XLSX.utils.sheet_to_json(sheet, { defval: null }), name);
      } catch {
        setToast({ message: 'Failed to parse Excel file', type: 'error' }); setLoading(false);
      }
    }, 300);
  }, [loadData]);

  const handleFileUpload = useCallback((e) => {
    const file = e.target?.files?.[0] || e.dataTransfer?.files?.[0];
    if (!file) return;
    const ext = file.name.toLowerCase().split('.').pop();
    if (!['csv', 'xlsx', 'xls'].includes(ext)) { setToast({ message: 'Please upload a .csv, .xlsx, or .xls file', type: 'error' }); return; }
    const reader = new FileReader();
    if (ext === 'csv') {
      reader.onload = (ev) => handleParseCSV(ev.target.result, file.name);
      reader.readAsText(file);
    } else {
      reader.onload = (ev) => handleParseXLSX(new Uint8Array(ev.target.result), file.name);
      reader.readAsArrayBuffer(file);
    }
  }, [handleParseCSV, handleParseXLSX]);

  const handleDrop = useCallback((e) => { e.preventDefault(); e.stopPropagation(); handleFileUpload(e); }, [handleFileUpload]);
  const handleSampleCSV = useCallback(() => handleParseCSV(generateSampleCSV(), 'sample_analytics.csv'), [handleParseCSV]);

  // ── Chart actions ──
  const handlePreview = () => {
    if (!csvData) { setToast({ message: 'Please upload a file first', type: 'error' }); return; }
    if (!xAxis) { setToast({ message: 'Please select an X-axis column', type: 'warning' }); return; }
    if (yAxis.length === 0) { setToast({ message: 'Please select at least one Y-axis column', type: 'warning' }); return; }
    setShowPreview(true);
  };

  const handleKeepChart = () => {
    const config = { 
      xAxis, 
      yAxis: [...yAxis], 
      chartType, 
      palette, 
      showTrendline, 
      legendLabels: { ...legendLabels }, 
      filterZeros, 
      topN, 
      logScale, 
      aggregation, 
      bubbleField,
      brushStartIndex,
      brushEndIndex
    };
    setSavedCharts(prev => [...prev, { config, id: Date.now() }]);
    setShowPreview(false);
    setToast({ message: 'Chart saved to report queue!', type: 'success' });
  };

  const handleDiscard = () => {
    setShowPreview(false);
    setYAxis([]);
    setChartType('bar');
    setShowTrendline(false);
    setLegendLabels({});
    setFilterZeros(true);
    setTopN('25');
    setLogScale(false);
    setAggregation('sum');
    setBubbleField('');
    setBrushStartIndex(0);
    setBrushEndIndex(24);
  };

  const handleRemoveChart = (id) => setSavedCharts(prev => prev.filter(c => c.id !== id));
  const handleMoveChart = (idx, dir) => {
    setSavedCharts(prev => {
      const arr = [...prev];
      const newIdx = idx + dir;
      if (newIdx < 0 || newIdx >= arr.length) return arr;
      [arr[idx], arr[newIdx]] = [arr[newIdx], arr[idx]];
      return arr;
    });
  };

  // ── Apply preset (Feature 4) ──
  const applyPreset = (preset) => {
    if (!csvData) { setToast({ message: 'Please upload a file first', type: 'error' }); return; }
    // Find matching column names
    const findCol = (candidates) => candidates.find(c => columns.includes(c)) || '';
    const xCol = findCol(preset.xAxis);
    const yCols = preset.yAxis.filter(c => columns.includes(c));
    if (!xCol || yCols.length === 0) {
      setToast({ message: `Preset requires columns: ${preset.xAxis[0]}, ${preset.yAxis.join(', ')} — not found in your data`, type: 'warning' });
      return;
    }
    setXAxis(xCol);
    setYAxis(yCols);
    setChartType(preset.chartType);
    setFilterZeros(preset.filterZeros);
    setTopN(preset.topN);
    setAggregation(preset.aggregation || 'sum');
    setShowPreview(true);
    setToast({ message: `Loaded preset: ${preset.name}`, type: 'success' });
  };

  // ── Enforce single Y for pie/donut ──
  useEffect(() => {
    if (PIE_TYPES.includes(chartType) && yAxis.length > 1) setYAxis([yAxis[0]]);
  }, [chartType]);

  const sameAxisWarning = xAxis && yAxis.includes(xAxis);
  const previewInsight = showPreview && visibleData.length > 0 ? generateInsight(yAxis, visibleData) : '';
  const currentConfig = { 
    xAxis, 
    yAxis, 
    chartType, 
    palette, 
    showTrendline, 
    legendLabels, 
    filterZeros, 
    topN, 
    logScale, 
    aggregation, 
    bubbleField,
    brushStartIndex,
    brushEndIndex
  };

  // Numeric columns for bubble field selection
  const numericColumns = useMemo(() => columns.filter(c => fieldMeta[c] === 'numeric' || fieldMeta[c] === 'ratio'), [columns, fieldMeta]);

  /* ═══ Theme classes ═══ */
  const bg = dark ? 'bg-[#0f172a]' : 'bg-slate-50';
  const panelBg = dark ? 'bg-slate-800/60' : 'bg-white';
  const panelBorder = dark ? 'border-slate-700/50' : 'border-slate-200';
  const textPrimary = dark ? 'text-white' : 'text-slate-800';
  const textSecondary = dark ? 'text-slate-400' : 'text-slate-500';
  const textMuted = dark ? 'text-slate-500' : 'text-slate-400';

  return (
    <div className={`h-screen flex flex-col ${bg} transition-colors duration-300 font-sans`}>
      {toast && <Toast {...toast} onClose={() => setToast(null)} />}
      {showReport && <ReportModal charts={savedCharts} data={csvData} fieldMeta={fieldMeta} onClose={() => setShowReport(false)} dark={dark} />}
      {showWeeklyReport && (
        <WeeklyReportModal 
          orgName={selectedOrg}
          targetDate={selectedWeeklyEndDate}
          weeklyReportData={weeklyReportData}
          selectedWeeklyMetrics={selectedWeeklyMetrics}
          weeklyChartType={weeklyChartType}
          onClose={() => setShowWeeklyReport(false)}
          dark={dark} 
        />
      )}

      {/* ═══════ HEADER ═══════ */}
      <header className={`no-print h-16 flex items-center justify-between px-6 border-b flex-shrink-0 ${panelBg} ${panelBorder} backdrop-blur-sm z-30`}
        style={{ borderColor: dark ? '#1e293b' : '#e2e8f0' }}>
        <div className="flex items-center gap-3">
          <button onClick={() => setSidebarOpen(!sidebarOpen)} className={`lg:hidden p-1.5 rounded-lg ${dark ? 'text-slate-400 hover:bg-slate-700' : 'text-slate-500 hover:bg-slate-100'}`}>
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 12h18M3 6h18M3 18h18" /></svg>
          </button>
          <LogoIcon />
          <div>
            <h1 className={`text-lg font-extrabold leading-tight ${textPrimary}`}>InsightForge</h1>
            <p className={`text-[10px] tracking-wide uppercase ${textMuted}`}>Turn your data into stories</p>
          </div>
        </div>

        {csvData && (
          <div className="hidden md:flex gap-1 bg-slate-200/40 dark:bg-slate-700/40 p-1 rounded-xl border border-slate-300/10 backdrop-blur-md">
            {/* Preserved for future changes:
            <button
              onClick={() => setViewMode('custom')}
              className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${
                viewMode === 'custom'
                  ? 'bg-blue-600 text-white shadow-md shadow-blue-600/15 scale-[1.02]'
                  : `${dark ? 'text-slate-300 hover:bg-slate-700/50 hover:text-white' : 'text-slate-600 hover:bg-slate-200/50 hover:text-slate-800'}`
              }`}
            >
              🎨 Custom Chart Builder
            </button>
            */}
            <button
              onClick={() => setViewMode('weekly')}
              className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-all ${
                viewMode === 'weekly'
                  ? 'bg-violet-600 text-white shadow-md shadow-violet-600/15 scale-[1.02]'
                  : `${dark ? 'text-slate-300 hover:bg-slate-700/50 hover:text-white' : 'text-slate-600 hover:bg-slate-200/50 hover:text-slate-800'}`
              }`}
            >
              ⚡ Weekly Org Report
            </button>
          </div>
        )}

        <div className="flex items-center gap-3">
          <button onClick={() => setDark(!dark)} className={`p-2 rounded-lg transition-all duration-300 ${dark ? 'bg-slate-700/80 text-amber-400 hover:bg-slate-600' : 'bg-slate-200 text-slate-600 hover:bg-slate-300'}`}
            title={dark ? 'Switch to light mode' : 'Switch to dark mode'}>
            {dark ? <SunIcon /> : <MoonIcon />}
          </button>
          <button 
            onClick={() => viewMode === 'custom' ? setShowReport(true) : setShowWeeklyReport(true)} 
            disabled={viewMode === 'custom' ? savedCharts.length === 0 : !weeklyReportData}
            className={`px-5 py-2 rounded-lg text-sm font-bold transition-all duration-300 ${
              (viewMode === 'custom' ? savedCharts.length === 0 : !weeklyReportData)
                ? (dark ? 'bg-slate-700 text-slate-500 cursor-not-allowed' : 'bg-slate-200 text-slate-400 cursor-not-allowed')
                : `bg-blue-600 text-white hover:bg-blue-500 shadow-lg shadow-blue-600/25`
            }`}>
            📊 Generate Report
          </button>
        </div>
      </header>

      {/* ═══════ BODY ═══════ */}
      <div className="flex flex-1 overflow-hidden no-print">
        {/* ═══ LEFT SIDEBAR ═══ */}
        <aside className={`${sidebarOpen ? 'translate-x-0' : '-translate-x-full'} lg:translate-x-0 fixed lg:static inset-y-16 left-0 z-20 w-[300px] flex-shrink-0 overflow-y-auto border-r transition-transform duration-300 ${panelBg} ${panelBorder}`}
          style={{ borderColor: dark ? '#1e293b' : '#e2e8f0' }}>
          {sidebarOpen && <div className="lg:hidden fixed inset-0 bg-black/50 -z-10" onClick={() => setSidebarOpen(false)} />}

          <div className="p-5 space-y-6">
            {/* ── Step 1: Upload ── */}
            <div>
              <h3 className={`text-xs font-bold uppercase tracking-widest mb-3 ${textMuted}`}>
                <span className="inline-flex items-center justify-center w-5 h-5 rounded-full bg-blue-600 text-white text-[10px] mr-2">1</span>
                Upload Data
              </h3>
              {!csvData ? (
                <div onDrop={handleDrop} onDragOver={e => e.preventDefault()} onClick={() => fileInputRef.current?.click()}
                  className={`rounded-xl border-2 border-dashed p-6 text-center cursor-pointer transition-all duration-300 group ${dark ? 'border-slate-600 hover:border-blue-500/50 hover:bg-blue-500/5' : 'border-slate-300 hover:border-blue-400 hover:bg-blue-50/50'}`}>
                  <UploadIcon />
                  <p className={`text-sm font-medium mb-1 ${dark ? 'text-slate-300' : 'text-slate-600'}`}>Drag & drop CSV or Excel file</p>
                  <p className={`text-xs ${textMuted}`}>.csv, .xlsx, .xls — click to browse</p>
                  <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFileUpload} className="hidden" />
                </div>
              ) : (
                <div className={`rounded-xl border p-4 ${dark ? 'border-slate-600/50 bg-slate-700/30' : 'border-slate-200 bg-slate-50'}`}>
                  <div className="flex items-center gap-2 mb-3">
                    <span className="w-2 h-2 rounded-full bg-emerald-400" />
                    <span className={`text-sm font-semibold truncate ${textPrimary}`}>{fileName}</span>
                  </div>
                  <p className={`text-xs mb-3 ${textMuted}`}>{rowCount} rows · {columns.length} columns</p>
                  <div className="overflow-x-auto rounded-lg border" style={{ borderColor: dark ? '#334155' : '#e2e8f0' }}>
                    <table className="w-full text-[10px]">
                      <thead>
                        <tr className={dark ? 'bg-slate-700/80' : 'bg-slate-100'}>
                          {columns.slice(0, 4).map(c => (
                            <th key={c} className={`px-2 py-1.5 text-left font-semibold font-mono-data truncate max-w-[70px] ${dark ? 'text-slate-300' : 'text-slate-600'}`}>{c}</th>
                          ))}
                          {columns.length > 4 && <th className={`px-2 py-1.5 ${textMuted}`}>+{columns.length - 4}</th>}
                        </tr>
                      </thead>
                      <tbody>
                        {previewRows.map((row, i) => (
                          <tr key={i} className={`border-t ${dark ? 'border-slate-700' : 'border-slate-200'}`}>
                            {columns.slice(0, 4).map(c => (
                              <td key={c} className={`px-2 py-1 font-mono-data truncate max-w-[70px] ${dark ? 'text-slate-400' : 'text-slate-500'}`}>{row[c] != null ? String(row[c]) : '—'}</td>
                            ))}
                            {columns.length > 4 && <td className={`px-2 py-1 ${textMuted}`}>…</td>}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <button onClick={() => { setCsvData(null); setColumns([]); setFieldMeta({}); setFileName(''); setRowCount(0); setPreviewRows([]); setShowPreview(false); setXAxis(''); setYAxis([]); }}
                    className={`mt-3 text-xs transition-colors ${dark ? 'text-slate-500 hover:text-red-400' : 'text-slate-400 hover:text-red-500'}`}
                  >
                    ✕ Remove file
                  </button>
                </div>
              )}
              {loading && (
                <div className="mt-3 space-y-2">
                  <div className={`h-3 rounded ${dark ? 'animate-shimmer' : 'animate-shimmer-light'}`} />
                  <div className={`h-3 rounded w-2/3 ${dark ? 'animate-shimmer' : 'animate-shimmer-light'}`} />
                </div>
              )}
              {!csvData && !loading && (
                <button onClick={handleSampleCSV} className={`mt-3 w-full text-xs py-2 rounded-lg border transition-colors ${dark ? 'border-slate-600 text-slate-400 hover:text-blue-400 hover:border-blue-500/50' : 'border-slate-300 text-slate-500 hover:text-blue-500 hover:border-blue-400'}`}>	
                  ✨ Generate sample data
                </button>
              )}
            </div>
            {viewMode === 'custom' && csvData && (
              <div>
                <h3 className={`text-xs font-bold uppercase tracking-widest mb-3 ${textMuted}`}>
                  <span className="inline-flex items-center justify-center w-5 h-5 rounded-full bg-violet-600 text-white text-[10px] mr-2">⚡</span>
                  Quick Presets
                </h3>
                <div className="space-y-1.5">
                  {PRESET_TEMPLATES.map(preset => (
                    <button key={preset.id} onClick={() => applyPreset(preset)}
                      className={`w-full text-left px-3 py-2.5 rounded-lg border text-xs transition-all hover:scale-[1.01] ${dark ? 'border-slate-700/50 bg-slate-800/30 hover:border-blue-500/40 hover:bg-blue-500/5 text-slate-300' : 'border-slate-200 bg-slate-50/80 hover:border-blue-400 hover:bg-blue-50 text-slate-600'}`}>
                      <div className="flex items-center gap-2">
                        <span className="text-base">{preset.icon}</span>
                        <div className="min-w-0">
                          <p className={`font-semibold truncate ${dark ? 'text-slate-200' : 'text-slate-700'}`}>{preset.name}</p>
                          <p className={`truncate ${textMuted}`}>{preset.description}</p>
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* ── Step 2: Configure Chart (Custom Mode) ── */}
            {viewMode === 'custom' && (
              <div className={!csvData ? 'opacity-40 pointer-events-none' : ''}>
                <h3 className={`text-xs font-bold uppercase tracking-widest mb-3 ${textMuted}`}>
                  <span className="inline-flex items-center justify-center w-5 h-5 rounded-full bg-blue-600 text-white text-[10px] mr-2">2</span>
                  Configure Chart
                </h3>

                <div className="space-y-4">
                  {/* X-Axis */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>X-Axis</label>
                    <SearchableDropdown options={columns} value={xAxis} onChange={setXAxis} placeholder="Select column…" dark={dark} />
                    {xAxis && <p className={`text-[10px] mt-0.5 ${textMuted}`}>{fieldMeta[xAxis] || 'auto'} field</p>}
                  </div>

                  {/* Y-Axis */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>
                      Y-Axis {PIE_TYPES.includes(chartType) && <span className="text-amber-400 font-normal">(single only for pie)</span>}
                    </label>
                    <SearchableDropdown options={columns} value={yAxis} onChange={setYAxis} placeholder="Select column(s)…" dark={dark} multiple maxSelections={PIE_TYPES.includes(chartType) ? 1 : 3} />
                  </div>

                  {sameAxisWarning && (
                    <p className="text-xs text-amber-400 bg-amber-500/10 px-3 py-2 rounded-lg">⚠️ X and Y axis are the same column</p>
                  )}

                  {/* ── Chart Type with Recommendations (Feature 1) ── */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Chart Type</label>
                    {recommendations.length > 0 ? (
                      <div className="space-y-1">
                        {recommendations.slice(0, 6).map((rec, idx) => {
                          const ct = CHART_TYPES.find(t => t.value === rec.type);
                          if (!ct) return null;
                          const isActive = chartType === rec.type;
                          const isRecommended = rec.rank <= 2;
                          return (
                            <button key={rec.type} onClick={() => setChartType(rec.type)} title={ct.tip}
                              className={`w-full text-left px-3 py-2 rounded-lg border text-xs flex items-center gap-2 transition-all ${
                                isActive
                                  ? 'border-blue-500 bg-blue-600/10 ring-1 ring-blue-500/30'
                                  : isRecommended
                                    ? (dark ? 'border-slate-600 bg-slate-800/40 hover:border-blue-500/50' : 'border-slate-200 bg-white hover:border-blue-400')
                                    : (dark ? 'border-slate-700/30 bg-slate-800/20 opacity-60 hover:opacity-100' : 'border-slate-100 bg-slate-50/50 opacity-50 hover:opacity-100')
                              }`}>
                              <span className="text-base flex-shrink-0">{ct.icon}</span>
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-1.5">
                                  <span className={`font-semibold ${isActive ? 'text-blue-400' : (dark ? 'text-slate-200' : 'text-slate-700')}`}>{ct.label}</span>
                                  {isRecommended && <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full bg-emerald-500/20 text-emerald-400">★ Best</span>}
                                </div>
                                <p className={`truncate ${textMuted}`}>{rec.reason}</p>
                              </div>
                              {isActive && <span className="w-2 h-2 rounded-full bg-blue-50 flex-shrink-0" />}
                            </button>
                          );
                        })}
                        {recommendations.length > 6 && (
                          <details className="mt-1">
                            <summary className={`text-xs cursor-pointer px-3 py-1 ${textMuted} hover:text-blue-400`}>More chart types…</summary>
                            <div className="mt-1 space-y-1">
                              {recommendations.slice(6).map(rec => {
                                const ct = CHART_TYPES.find(t => t.value === rec.type);
                                if (!ct) return null;
                                const isActive = chartType === rec.type;
                                return (
                                  <button key={rec.type} onClick={() => setChartType(rec.type)} title={ct.tip}
                                    className={`w-full text-left px-3 py-1.5 rounded-lg border text-xs flex items-center gap-2 transition-all ${
                                      isActive ? 'border-blue-500 bg-blue-600/10' : (dark ? 'border-slate-700/30 opacity-60 hover:opacity-100' : 'border-slate-100 opacity-50 hover:opacity-100')
                                    }`}>
                                    <span>{ct.icon}</span>
                                    <span className={isActive ? 'text-blue-400 font-semibold' : (dark ? 'text-slate-300' : 'text-slate-600')}>{ct.label}</span>
                                  </button>
                                );
                              })}
                            </div>
                          </details>
                        )}
                      </div>
                    ) : (
                      <SearchableDropdown
                        options={CHART_TYPES.map(t => t.label)}
                        value={CHART_TYPES.find(t => t.value === chartType)?.label || ''}
                        onChange={(label) => { const ct = CHART_TYPES.find(t => t.label === label); if (ct) setChartType(ct.value); }}
                        placeholder="Select chart type…" dark={dark} />
                    )}
                  </div>

                  {/* Aggregation toggle (Feature 3) */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Aggregation</label>
                    <div className="flex gap-1.5">
                      {['sum', 'avg', 'count'].map(mode => (
                        <button key={mode} onClick={() => setAggregation(mode)}
                          className={`flex-1 py-1.5 rounded-lg text-xs font-semibold capitalize transition-all ${
                            aggregation === mode ? 'bg-blue-600 text-white shadow-sm'
                            : (dark ? 'bg-slate-700/60 text-slate-400 hover:bg-slate-700' : 'bg-white text-slate-500 hover:bg-slate-100 border border-slate-200')
                          }`}>
                          {mode === 'avg' ? 'Average' : mode === 'count' ? 'Count' : 'Sum'}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Bubble field selector */}
                  {chartType === 'bubble' && (
                    <div>
                      <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Bubble Size Field</label>
                      <SearchableDropdown options={numericColumns} value={bubbleField} onChange={setBubbleField} placeholder="Select size metric…" dark={dark} />
                    </div>
                  )}

                  {/* Color palette */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Color Palette</label>
                    <div className="flex gap-2 flex-wrap">
                      {Object.entries(COLOR_PALETTES).map(([name, clrs]) => (
                        <button key={name} onClick={() => setPalette(name)}
                          className={`flex gap-0.5 p-1.5 rounded-lg border transition-all ${palette === name ? 'border-blue-500 ring-1 ring-blue-500/30' : (dark ? 'border-slate-600 hover:border-slate-500' : 'border-slate-300 hover:border-slate-400')}`} title={name}>
                          {clrs.slice(0, 3).map((c, i) => <span key={i} className="w-3 h-3 rounded-sm" style={{ backgroundColor: c }} />)}
                        </button>
                      ))}
                    </div>
                    <p className={`text-[10px] mt-1 ${textMuted}`}>{palette}</p>
                  </div>

                  {/* Data filters */}
                  <div className={`rounded-lg border p-3 space-y-3 ${dark ? 'border-slate-700/50 bg-slate-800/30' : 'border-slate-200 bg-slate-50'}`}>
                    <p className={`text-[10px] font-bold uppercase tracking-wider ${textMuted}`}>Data Filters</p>
                    <label className="flex items-center gap-2.5 cursor-pointer">
                      <div className="relative">
                        <input type="checkbox" checked={filterZeros} onChange={e => setFilterZeros(e.target.checked)} className="sr-only peer" />
                        <div className={`w-9 h-5 rounded-full transition-colors ${dark ? 'bg-slate-600 peer-checked:bg-emerald-600' : 'bg-slate-300 peer-checked:bg-emerald-500'}`} />
                        <div className="absolute top-0.5 left-0.5 w-4 h-4 bg-white rounded-full transition-transform peer-checked:translate-x-4 shadow" />
                      </div>
                      <span className={`text-sm ${dark ? 'text-slate-300' : 'text-slate-600'}`}>Hide zero values</span>
                    </label>
                    <div>
                      <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Show Top</label>
                      <div className="flex gap-1.5">
                        {['10', '15', '25', '50', 'all'].map(opt => (
                          <button key={opt} onClick={() => setTopN(opt)}
                            className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                              topN === opt ? 'bg-blue-600 text-white shadow-sm' : (dark ? 'bg-slate-700/60 text-slate-400 hover:bg-slate-700' : 'bg-white text-slate-500 hover:bg-slate-100 border border-slate-200')
                            }`}>
                            {opt === 'all' ? 'All' : opt}
                          </button>
                        ))}
                      </div>
                    </div>
                    {!PIE_TYPES.includes(chartType) && (
                      <label className="flex items-center gap-2.5 cursor-pointer">
                        <div className="relative">
                          <input type="checkbox" checked={logScale} onChange={e => setLogScale(e.target.checked)} className="sr-only peer" />
                          <div className={`w-9 h-5 rounded-full transition-colors ${dark ? 'bg-slate-600 peer-checked:bg-violet-600' : 'bg-slate-300 peer-checked:bg-violet-500'}`} />
                          <div className="absolute top-0.5 left-0.5 w-4 h-4 bg-white rounded-full transition-transform peer-checked:translate-x-4 shadow" />
                        </div>
                        <span className={`text-sm ${dark ? 'text-slate-300' : 'text-slate-600'}`}>Log scale</span>
                      </label>
                    )}
                  </div>

                  {!PIE_TYPES.includes(chartType) && (
                    <label className="flex items-center gap-2.5 cursor-pointer group">
                      <div className="relative">
                        <input type="checkbox" checked={showTrendline} onChange={e => setShowTrendline(e.target.checked)} className="sr-only peer" />
                        <div className={`w-9 h-5 rounded-full transition-colors ${dark ? 'bg-slate-600 peer-checked:bg-blue-600' : 'bg-slate-300 peer-checked:bg-blue-500'}`} />
                        <div className="absolute top-0.5 left-0.5 w-4 h-4 bg-white rounded-full transition-transform peer-checked:translate-x-4 shadow" />
                      </div>
                      <span className={`text-sm ${dark ? 'text-slate-300' : 'text-slate-600'}`}>Add trendline</span>
                    </label>
                  )}

                  {yAxis.length > 0 && (
                    <div>
                      <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Legend Labels</label>
                      <div className="space-y-2">
                        {yAxis.map(col => (
                          <div key={col} className="flex items-center gap-2">
                            <span className={`text-[10px] font-mono-data w-16 truncate ${textMuted}`}>{col}</span>
                            <input type="text" value={legendLabels[col] || ''} onChange={e => setLegendLabels(prev => ({ ...prev, [col]: e.target.value }))} placeholder={col}
                              className={`flex-1 px-2.5 py-1.5 rounded-lg border text-sm transition-colors focus:outline-none focus:ring-1 focus:ring-blue-500/50 ${dark ? 'bg-slate-800/80 border-slate-600 text-slate-200 placeholder-slate-500' : 'bg-white border-slate-300 text-slate-700 placeholder-slate-400'}`} />
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <button onClick={handlePreview}
                    className="w-full py-2.5 rounded-lg bg-blue-600 text-white text-sm font-bold hover:bg-blue-500 transition-all shadow-lg shadow-blue-600/20 active:scale-[0.98]">
                    👁️ Preview Chart
                  </button>
                </div>
              </div>
            )}

            {/* ── Step 2: Configure Weekly Report (Weekly Mode) ── */}
            {viewMode === 'weekly' && (
              <div className={!csvData ? 'opacity-40 pointer-events-none' : ''}>
                <h3 className={`text-xs font-bold uppercase tracking-widest mb-3 ${textMuted}`}>
                  <span className="inline-flex items-center justify-center w-5 h-5 rounded-full bg-violet-600 text-white text-[10px] mr-2">2</span>
                  Configure Report
                </h3>

                <div className="space-y-4">
                  {/* Organization Selector */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Select Organization</label>
                    <SearchableDropdown 
                      options={uniqueOrgs} 
                      value={selectedOrg} 
                      onChange={setSelectedOrg} 
                      placeholder="Select organization…" 
                      dark={dark} 
                    />
                  </div>

                  {/* Date Range Selector */}
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Start Date</label>
                      <SearchableDropdown 
                        options={uniqueOrgDates} 
                        value={selectedWeeklyStartDate} 
                        onChange={(date) => {
                          setSelectedWeeklyStartDate(date);
                          if (weeklyHistoryLimit === 'mau') {
                            const startIdx = uniqueOrgDates.indexOf(date);
                            if (startIdx !== -1) {
                              const targetEndIdx = Math.min(uniqueOrgDates.length - 1, startIdx + 3);
                              setSelectedWeeklyEndDate(uniqueOrgDates[targetEndIdx]);
                            }
                          }
                        }} 
                        placeholder="Start date…" 
                        dark={dark} 
                      />
                    </div>
                    <div>
                      <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>End Date</label>
                      <SearchableDropdown 
                        options={uniqueOrgDates} 
                        value={selectedWeeklyEndDate} 
                        onChange={(date) => {
                          setSelectedWeeklyEndDate(date);
                          if (weeklyHistoryLimit === 'mau') {
                            const endIdx = uniqueOrgDates.indexOf(date);
                            if (endIdx !== -1) {
                              const targetStartIdx = Math.max(0, endIdx - 3);
                              setSelectedWeeklyStartDate(uniqueOrgDates[targetStartIdx]);
                            }
                          }
                        }} 
                        placeholder="End date…" 
                        dark={dark} 
                      />
                    </div>
                  </div>

                  {/* Display Metrics */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Display Metrics</label>
                    <SearchableDropdown 
                      options={columns} 
                      value={selectedWeeklyMetrics} 
                      onChange={setSelectedWeeklyMetrics} 
                      placeholder="Select columns…" 
                      dark={dark} 
                      multiple
                    />
                  </div>

                  {/* Chart Type */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>Chart Type</label>
                    <div className="flex gap-1.5">
                      {[
                        { value: 'line', label: 'Line', icon: '📈' },
                        { value: 'bar', label: 'Bar', icon: '📊' },
                        { value: 'area', label: 'Area', icon: '📉' }
                      ].map(opt => (
                        <button key={opt.value} onClick={() => setWeeklyChartType(opt.value)}
                          className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition-all flex items-center justify-center gap-1 ${
                            weeklyChartType === opt.value 
                              ? 'bg-violet-600 text-white shadow-sm' 
                              : (dark ? 'bg-slate-700/60 text-slate-400 hover:bg-slate-700' : 'bg-white text-slate-500 hover:bg-slate-100 border border-slate-200')
                          }`}>
                          <span>{opt.icon}</span>
                          <span>{opt.label}</span>
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* History Duration */}
                  <div>
                    <label className={`block text-xs font-semibold mb-1.5 ${textSecondary}`}>History Duration</label>
                    <div className="flex gap-1.5">
                      {[
                        { value: 'mau', label: 'MAU (4W)' },
                        { value: 'quarter', label: 'Quarter (12W)' }
                      ].map(opt => (
                        <button key={opt.value} onClick={() => {
                          setWeeklyHistoryLimit(opt.value);
                          if (opt.value === 'mau') {
                            const startIdx = uniqueOrgDates.indexOf(selectedWeeklyStartDate);
                            if (startIdx !== -1) {
                              const targetEndIdx = Math.min(uniqueOrgDates.length - 1, startIdx + 3);
                              setSelectedWeeklyEndDate(uniqueOrgDates[targetEndIdx]);
                            }
                          } else if (opt.value === 'quarter') {
                            if (uniqueOrgDates.length > 0) {
                              setSelectedWeeklyStartDate(uniqueOrgDates[0]);
                              setSelectedWeeklyEndDate(uniqueOrgDates[uniqueOrgDates.length - 1]);
                            }
                          }
                        }}
                          className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                            weeklyHistoryLimit === opt.value 
                              ? 'bg-violet-600 text-white shadow-sm' 
                              : (dark ? 'bg-slate-700/60 text-slate-400 hover:bg-slate-700' : 'bg-white text-slate-500 hover:bg-slate-100 border border-slate-200')
                          }`}>
                          {opt.label}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </aside>

        {/* ═══ MAIN CANVAS ═══ */}
        <main className={`flex-1 overflow-y-auto ${dark ? 'canvas-grid-dark' : 'canvas-grid-light'}`}>
          <div className="p-8 max-w-5xl mx-auto">
            {viewMode === 'custom' ? (
              !showPreview ? (
                <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-center animate-fadeIn">
                  <svg width="100" height="100" viewBox="0 0 100 100" fill="none" className={`mb-6 ${dark ? 'text-slate-700' : 'text-slate-300'}`}>
                    <rect x="10" y="45" width="15" height="40" rx="3" fill="currentColor" opacity="0.5" />
                    <rect x="30" y="30" width="15" height="55" rx="3" fill="currentColor" opacity="0.65" />
                    <rect x="50" y="20" width="15" height="65" rx="3" fill="currentColor" opacity="0.8" />
                    <rect x="70" y="10" width="15" height="75" rx="3" fill="currentColor" opacity="0.95" />
                    <path d="M15 43L37 28L57 18L77 8" stroke="currentColor" strokeWidth="2" strokeDasharray="4 2" opacity="0.4" />
                  </svg>
                  <h2 className={`text-xl font-bold mb-2 ${dark ? 'text-slate-400' : 'text-slate-500'}`}>No chart preview yet</h2>
                  <p className={`text-sm max-w-xs ${textMuted}`}>
                    Upload a CSV, configure your axes, and click <span className="font-semibold text-blue-400">"Preview Chart"</span> to visualize your data.
                  </p>
                  {csvData && (
                    <p className={`text-xs mt-4 ${textMuted}`}>
                      💡 Try a <span className="font-semibold text-violet-400">Quick Preset</span> in the sidebar for instant charts
                    </p>
                  )}
                </div>
              ) : (
                <div className="animate-fadeIn">
                  {/* Chart title */}
                  <div className="flex items-center justify-between mb-4">
                    <div>
                      <h2 className={`text-lg font-bold ${dark ? 'text-white' : 'text-slate-800'}`}>
                        {yAxis.join(', ')} <span className={`font-normal ${dark ? 'text-slate-500' : 'text-slate-400'}`}>vs</span> {xAxis}
                      </h2>
                      <p className={`text-xs ${dark ? 'text-slate-500' : 'text-slate-400'}`}>
                        {CHART_TYPES.find(t => t.value === chartType)?.label} · {csvData.length} data points · Aggregation: {aggregation}
                      </p>
                    </div>
                    <span className={`text-[10px] font-mono-data px-3 py-1.5 rounded-lg ${dark ? 'bg-slate-800 text-slate-500' : 'bg-slate-100 text-slate-400'}`}>
                      {palette} palette
                    </span>
                  </div>

                  <StatsRibbon yColumns={yAxis} data={visibleData} dark={dark} palette={palette} />

                  <div className={`rounded-2xl border p-5 mb-6 ${dark ? 'bg-slate-800/40 border-slate-700/50' : 'bg-white border-slate-200 shadow-sm'}`}>
                    <ChartRenderer 
                      config={currentConfig} 
                      data={csvData} 
                      fieldMeta={fieldMeta} 
                      dark={dark} 
                      onBrushChange={(e) => { 
                        if (e) { 
                          setBrushStartIndex(e.startIndex); 
                          setBrushEndIndex(e.endIndex); 
                        } 
                      }} 
                    />
                  </div>

                  <div className={`rounded-xl border px-5 py-4 mb-6 animate-slideUp ${dark ? 'bg-slate-800/30 border-slate-700/50' : 'bg-blue-50/70 border-blue-100'}`}>
                    <p className={`text-sm leading-relaxed ${dark ? 'text-slate-300' : 'text-slate-600'}`}>💡 {previewInsight}</p>
                  </div>

                  <div className="flex gap-3 animate-slideUp" style={{ animationDelay: '0.1s' }}>
                    <button onClick={handleKeepChart}
                      className="flex-1 py-3 rounded-xl bg-emerald-600 text-white font-bold text-sm hover:bg-emerald-500 transition-all shadow-lg shadow-emerald-600/20 active:scale-[0.98]">
                      ✅ Keep this chart
                    </button>
                    <button onClick={handleDiscard}
                      className={`flex-1 py-3 rounded-xl font-bold text-sm transition-all border active:scale-[0.98] ${dark ? 'border-slate-600 text-slate-400 hover:bg-slate-800 hover:text-red-400' : 'border-slate-300 text-slate-500 hover:bg-red-50 hover:text-red-500'}`}>
                      ❌ Discard
                    </button>
                  </div>
                </div>
              )
            ) : (
              // Weekly Org Report Dashboard
              !csvData ? (
                <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-center animate-fadeIn">
                  <svg width="100" height="100" viewBox="0 0 100 100" fill="none" className={`mb-6 ${dark ? 'text-slate-700' : 'text-slate-300'}`}>
                    <rect x="10" y="45" width="15" height="40" rx="3" fill="currentColor" opacity="0.5" />
                    <rect x="30" y="30" width="15" height="55" rx="3" fill="currentColor" opacity="0.65" />
                    <rect x="50" y="20" width="15" height="65" rx="3" fill="currentColor" opacity="0.8" />
                    <rect x="70" y="10" width="15" height="75" rx="3" fill="currentColor" opacity="0.95" />
                    <path d="M15 43L37 28L57 18L77 8" stroke="currentColor" strokeWidth="2" strokeDasharray="4 2" opacity="0.4" />
                  </svg>
                  <h2 className={`text-xl font-bold mb-2 ${dark ? 'text-slate-400' : 'text-slate-500'}`}>No weekly report data</h2>
                  <p className={`text-sm max-w-xs ${textMuted}`}>
                    Upload a CSV or Excel file to generate the Weekly Organization Report.
                  </p>
                </div>
              ) : !weeklyReportData ? (
                <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-center animate-fadeIn">
                  <h2 className={`text-xl font-bold mb-2 ${dark ? 'text-slate-400' : 'text-slate-500'}`}>Weekly Report Configuration Incomplete</h2>
                  <p className={`text-sm max-w-xs ${textMuted}`}>
                    Please select an organization and target date in the sidebar to generate the weekly report.
                  </p>
                </div>
              ) : (
                <div className="animate-fadeIn">
                  {/* Metadata banner */}
                  {(() => {
                    const firstRow = weeklyReportData.historyRows[0] || {};
                    const regionKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('region')) || 'Focus Region';
                    const cohortKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('cohort')) || 'India Cohort';
                    const region = firstRow[regionKey] || '—';
                    const cohort = firstRow[cohortKey] || '—';
                    
                    return (
                      <div className="mb-6 flex flex-col md:flex-row md:items-center justify-between gap-4 border-b pb-5" style={{ borderColor: dark ? '#334155' : '#e2e8f0' }}>
                        <div>
                          <h2 className={`text-2xl font-extrabold tracking-tight ${textPrimary}`}>{selectedOrg}</h2>
                          <p className={`text-xs mt-1 ${textSecondary}`}>
                            Range: <span className="font-semibold text-violet-500">{selectedWeeklyStartDate}</span> to <span className="font-semibold text-violet-500">{selectedWeeklyEndDate}</span>
                          </p>
                        </div>
                        <div className="flex gap-2">
                          <span className={`text-[10px] font-bold uppercase tracking-wider px-2.5 py-1 rounded-lg ${dark ? 'bg-slate-800 text-slate-400 border border-slate-700/50' : 'bg-slate-100 text-slate-600 border border-slate-200'}`}>
                            Region: {region}
                          </span>
                          <span className={`text-[10px] font-bold uppercase tracking-wider px-2.5 py-1 rounded-lg ${dark ? 'bg-violet-950/40 text-violet-400 border border-violet-800/30' : 'bg-violet-50 text-violet-600 border border-violet-100'}`}>
                            Cohort: {cohort}
                          </span>
                        </div>
                      </div>
                    );
                  })()}

                  {/* Top Stats Cards Grid */}
                  {(() => {
                    const colors = [
                      { text: 'text-blue-500', hex: '#3b82f6' },
                      { text: 'text-violet-500', hex: '#8b5cf6' },
                      { text: 'text-emerald-500', hex: '#10b981' },
                      { text: 'text-amber-500', hex: '#f59e0b' },
                      { text: 'text-rose-500', hex: '#f43f5e' },
                      { text: 'text-cyan-500', hex: '#06b6d4' },
                      { text: 'text-pink-500', hex: '#ec4899' },
                      { text: 'text-indigo-500', hex: '#6366f1' },
                    ];
                    const metricsList = selectedWeeklyMetrics.map((colName, idx) => {
                      const color = colors[idx % colors.length];
                      return {
                        key: colName,
                        label: colName,
                        colorClass: color.text,
                        colorHex: color.hex
                      };
                    });

                    return (
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
                        {metricsList.map(m => {
                          const colName = m.key;
                          const currentVal = Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx][colName]) || 0;
                          const prevVal = weeklyReportData.selectedIdx > 0 
                            ? (Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx - 1][colName]) || 0) 
                            : null;
                          const changePct = (prevVal !== null && prevVal > 0)
                            ? ((currentVal - prevVal) / prevVal) * 100
                            : null;

                          return (
                            <div key={m.key} className={`rounded-xl border p-4 ${dark ? 'bg-slate-800/40 border-slate-700/50' : 'bg-white border-slate-200 shadow-sm'}`}>
                              <p className={`text-[10px] font-bold uppercase tracking-wider mb-1 ${textMuted}`}>{m.label}</p>
                              <p className={`text-2xl font-extrabold ${m.colorClass}`}>{currentVal.toLocaleString()}</p>
                              <div className="flex items-center gap-1.5 mt-1.5">
                                {changePct !== null ? (
                                  <>
                                    <span className={`text-xs font-bold ${changePct >= 0 ? 'text-emerald-500' : 'text-red-500'}`}>
                                      {changePct >= 0 ? '↑' : '↓'} {Math.abs(changePct).toFixed(1)}%
                                    </span>
                                    <span className={`text-[10px] ${textMuted}`}>vs last week</span>
                                  </>
                                ) : (
                                  <span className={`text-[10px] ${textMuted}`}>No prior data</span>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })()}

                  {/* Historical Trend Line Graph */}
                  {(() => {
                    const colors = [
                      { text: 'text-blue-500', hex: '#3b82f6' },
                      { text: 'text-violet-500', hex: '#8b5cf6' },
                      { text: 'text-emerald-500', hex: '#10b981' },
                      { text: 'text-amber-500', hex: '#f59e0b' },
                      { text: 'text-rose-500', hex: '#f43f5e' },
                      { text: 'text-cyan-500', hex: '#06b6d4' },
                      { text: 'text-pink-500', hex: '#ec4899' },
                      { text: 'text-indigo-500', hex: '#6366f1' },
                    ];
                    const activeMetrics = selectedWeeklyMetrics.map((colName, idx) => {
                      const color = colors[idx % colors.length];
                      return {
                        key: colName,
                        label: colName,
                        colorClass: color.text,
                        colorHex: color.hex
                      };
                    });

                    return (
                      <div className={`rounded-2xl border p-5 mb-6 ${dark ? 'bg-slate-800/40 border-slate-700/50' : 'bg-white border-slate-200 shadow-sm'}`}>
                        <div className="flex items-center justify-between mb-4">
                          <div>
                            <h3 className={`text-sm font-bold ${textPrimary}`}>Historical Trend Progression</h3>
                            <p className={`text-xs ${textMuted}`}>Displaying chronological weekly metrics with target week indicator</p>
                          </div>
                        </div>
                        <div style={{ width: '100%', height: 350 }}>
                          <ResponsiveContainer>
                            <ComposedChart data={weeklyReportData.historyRows}>
                              <CartesianGrid strokeDasharray="3 3" stroke={dark ? '#334155' : '#e2e8f0'} vertical={false} />
                              <XAxis 
                                dataKey={weeklyReportData.dateCol} 
                                stroke={dark ? '#64748b' : '#94a3b8'} 
                                fontSize={10} 
                                tickLine={false} 
                                axisLine={false}
                                tickFormatter={(str) => {
                                  const d = new Date(str);
                                  return isNaN(d) ? str : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                                }}
                              />
                              <YAxis 
                                stroke={dark ? '#64748b' : '#94a3b8'} 
                                fontSize={10} 
                                tickLine={false} 
                                axisLine={false}
                                tickFormatter={(val) => val.toLocaleString()}
                              />
                              <Tooltip content={<CustomTooltip dark={dark} />} />
                              <Legend iconType="circle" wrapperStyle={{ fontSize: 11, paddingTop: 10 }} />
                              
                              <ReferenceLine 
                                x={selectedWeeklyEndDate} 
                                stroke={dark ? '#a855f7' : '#8b5cf6'} 
                                strokeWidth={2}
                                strokeDasharray="4 4"
                                label={{ value: 'Target Week', fill: dark ? '#c084fc' : '#6b21a8', fontSize: 10, position: 'top' }} 
                              />
                              
                              {activeMetrics.map(m => {
                                const colName = m.key;
                                const color = m.colorHex;
                                if (weeklyChartType === 'bar') {
                                  return (
                                    <Bar 
                                      key={m.key} 
                                      dataKey={colName} 
                                      name={m.label} 
                                      fill={color} 
                                      radius={[4, 4, 0, 0]}
                                    />
                                  );
                                } else if (weeklyChartType === 'area') {
                                  return (
                                    <Area 
                                      key={m.key} 
                                      type="monotone"
                                      dataKey={colName} 
                                      name={m.label} 
                                      stroke={color} 
                                      fill={color} 
                                      fillOpacity={0.15}
                                      strokeWidth={2}
                                    />
                                  );
                                } else {
                                  return (
                                    <Line 
                                      key={m.key} 
                                      type="monotone" 
                                      dataKey={colName} 
                                      name={m.label} 
                                      stroke={color} 
                                      strokeWidth={3} 
                                      dot={{ r: 4, strokeWidth: 1 }}
                                      activeDot={{ r: 7, strokeWidth: 0 }}
                                    />
                                  );
                                }
                              })}
                            </ComposedChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                    );
                  })()}

                  {/* Rolling Details Panel */}
                  <div className="mb-6">
                    <h3 className={`text-sm font-bold mb-3.5 ${textPrimary}`}>Rolling {weeklyReportData.windowSize}-Week Active Window</h3>
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                      {weeklyReportData.activeWindow.map((week, idx) => {
                        const isSelectedWeek = week.date === selectedWeeklyEndDate;
                        const colors = [
                          { text: 'text-blue-500', hex: '#3b82f6' },
                          { text: 'text-violet-500', hex: '#8b5cf6' },
                          { text: 'text-emerald-500', hex: '#10b981' },
                          { text: 'text-amber-500', hex: '#f59e0b' },
                          { text: 'text-rose-500', hex: '#f43f5e' },
                          { text: 'text-cyan-500', hex: '#06b6d4' },
                          { text: 'text-pink-500', hex: '#ec4899' },
                          { text: 'text-indigo-500', hex: '#6366f1' },
                        ];
                        const activeMetrics = selectedWeeklyMetrics.map((colName, idx) => {
                          const color = colors[idx % colors.length];
                          return {
                            key: colName,
                            label: colName,
                            colorClass: color.text,
                            colorHex: color.hex
                          };
                        });

                        return (
                          <div 
                            key={week.label} 
                            className={`rounded-xl border p-4 transition-all ${
                              isSelectedWeek 
                                ? 'bg-violet-600/5 border-violet-500 ring-1 ring-violet-500/20 shadow-md shadow-violet-500/5' 
                                : (dark ? 'bg-slate-800/40 border-slate-700/50 hover:bg-slate-800/60' : 'bg-white border-slate-200 hover:shadow-sm')
                            }`}
                          >
                            <div className="flex items-center justify-between mb-2">
                              <span className={`text-[10px] font-extrabold px-2 py-0.5 rounded-full ${
                                isSelectedWeek 
                                  ? 'bg-violet-500 text-white' 
                                  : (dark ? 'bg-slate-700 text-slate-300' : 'bg-slate-100 text-slate-600')
                              }`}>
                                {week.label} {isSelectedWeek && ' (Target)'}
                              </span>
                              <span className={`text-[10px] font-mono-data ${textMuted}`}>
                                {new Date(week.date).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
                              </span>
                            </div>
                            
                            <div className="space-y-1.5 mt-3">
                              {activeMetrics.map(m => {
                                const colName = m.key;
                                const val = week.rawRow[colName] != null ? Number(week.rawRow[colName]) : null;
                                const color = `${m.colorClass} font-semibold`;
                                
                                return (
                                  <div key={m.key} className="flex justify-between items-center text-xs">
                                    <span className={textMuted}>{m.label.replace('Current ', '')}</span>
                                    <span className={color}>
                                      {val !== null ? val.toLocaleString() : '—'}
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Transition Card (Visual Added/Dropped Animation/Explainer) */}
                  <div className={`rounded-xl border p-5 ${dark ? 'bg-slate-800/30 border-slate-700/50' : 'bg-slate-100/50 border-slate-200'}`}>
                    <h3 className={`text-sm font-bold mb-3.5 ${textPrimary}`}>Rolling Window Progression ({weeklyReportData.windowSize}-Week Shift)</h3>
                    <div className="flex flex-col lg:flex-row items-center justify-between gap-6">
                      <div className="flex-1 min-w-[280px]">
                        <p className={`text-sm leading-relaxed ${dark ? 'text-slate-300' : 'text-slate-600'}`}>
                          Active users are computed over a rolling {weeklyReportData.windowSize}-week duration. 
                          When progressing to week <span className="font-semibold text-violet-500 font-mono">{selectedWeeklyEndDate}</span>:
                        </p>
                        <ul className={`list-disc pl-5 mt-2 space-y-1 text-xs ${textSecondary}`}>
                          {weeklyReportData.droppedWeekDate ? (
                            <li>
                              The oldest week <span className="text-red-500 font-semibold font-mono">{weeklyReportData.droppedWeekDate}</span> is <span className="text-red-500 font-semibold">dropped</span> from the rolling set.
                            </li>
                          ) : (
                            <li>Initial weeks: no week dropped yet as database window is building up.</li>
                          )}
                          <li>
                            The new week <span className="text-emerald-500 font-semibold font-mono">{selectedWeeklyEndDate}</span> is <span className="text-emerald-500 font-semibold">added</span>.
                          </li>
                        </ul>
                      </div>

                      <div className="flex items-center gap-3 md:gap-4 flex-shrink-0">
                        {weeklyReportData.droppedWeekDate ? (
                          <div className={`w-36 rounded-xl border p-3.5 text-center ${dark ? 'bg-red-500/5 border-red-500/20' : 'bg-red-50/50 border-red-100'}`}>
                            <span className="text-[9px] font-extrabold uppercase tracking-wider px-2 py-0.5 rounded-full bg-red-500/10 text-red-500">
                              Dropped
                            </span>
                            <p className={`text-xs font-mono font-bold mt-2 ${textPrimary}`}>{weeklyReportData.droppedWeekDate}</p>
                            {(() => {
                              const firstMetricCol = selectedWeeklyMetrics[0];
                              const val = firstMetricCol ? (Number(weeklyReportData.droppedWeekRow?.[firstMetricCol]) || 0) : 0;
                              return (
                                <p className="text-sm font-extrabold text-red-500/80 mt-1">
                                  -{val.toLocaleString()}
                                </p>
                              );
                            })()}
                          </div>
                        ) : (
                          <div className={`w-36 rounded-xl border p-3.5 text-center border-dashed ${dark ? 'border-slate-700 bg-slate-800/10 text-slate-600' : 'border-slate-200 bg-slate-50 text-slate-400'}`}>
                            <span className="text-[9px] font-bold">No Drop</span>
                            <p className="text-xs mt-2 font-mono">—</p>
                          </div>
                        )}

                        <div className="flex flex-col items-center">
                          <span className="text-xl md:text-2xl text-violet-500">➔</span>
                          <span className={`text-[8px] font-semibold tracking-wide uppercase ${textMuted}`}>Shift</span>
                        </div>

                        <div className={`w-36 rounded-xl border p-3.5 text-center ${dark ? 'bg-emerald-500/5 border-emerald-500/20' : 'bg-emerald-50/50 border-emerald-100'}`}>
                          <span className="text-[9px] font-extrabold uppercase tracking-wider px-2 py-0.5 rounded-full bg-emerald-500/10 text-emerald-500">
                            Added
                          </span>
                          <p className={`text-xs font-mono font-bold mt-2 ${textPrimary}`}>{selectedWeeklyEndDate}</p>
                          {(() => {
                            const firstMetricCol = selectedWeeklyMetrics[0];
                            const val = firstMetricCol ? (Number(weeklyReportData.orgRows[weeklyReportData.selectedIdx]?.[firstMetricCol]) || 0) : 0;
                            return (
                              <p className="text-sm font-extrabold text-emerald-500/80 mt-1">
                                +{val.toLocaleString()}
                              </p>
                            );
                          })()}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              )
            )}
          </div>
        </main>

        {/* ═══ RIGHT PANEL — REPORT QUEUE ═══ */}
        {viewMode === 'custom' && (
          <aside className={`hidden xl:block w-[280px] flex-shrink-0 overflow-y-auto border-l ${panelBg}`}
            style={{ borderColor: dark ? '#1e293b' : '#e2e8f0' }}>
            <div className="p-5">
              <h3 className={`text-xs font-bold uppercase tracking-widest mb-4 ${textMuted}`}>
                Your Report ({savedCharts.length} chart{savedCharts.length !== 1 ? 's' : ''})
              </h3>
              {savedCharts.length === 0 ? (
                <div className={`text-center py-10 ${textMuted}`}>
                  <p className="text-sm mb-1">Add your first chart →</p>
                  <p className="text-xs">Preview and keep a chart to add it here</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {savedCharts.map((chart, idx) => {
                    const processedChartData = processChartData(csvData, chart.config, fieldMeta);
                    const visibleChartData = chart.config.brushStartIndex !== undefined && chart.config.brushEndIndex !== undefined
                      ? processedChartData.slice(chart.config.brushStartIndex, chart.config.brushEndIndex + 1)
                      : processedChartData;
                    const insight = generateInsight(chart.config.yAxis, visibleChartData);
                    const typeLabel = CHART_TYPES.find(t => t.value === chart.config.chartType)?.label;
                    return (
                      <div key={chart.id} className={`rounded-xl border p-3 animate-slideUp ${dark ? 'bg-slate-700/40 border-slate-600/50' : 'bg-slate-50 border-slate-200'}`}>
                        <div className={`rounded-lg overflow-hidden mb-2.5 ${dark ? 'bg-slate-800/60' : 'bg-white'}`}>
                          <ChartRenderer config={chart.config} data={csvData} fieldMeta={fieldMeta} height={120} dark={dark} animate={false} />
                        </div>
                        <div className="flex items-start justify-between gap-2 mb-1.5">
                          <div className="min-w-0">
                            <p className={`text-xs font-semibold truncate ${textPrimary}`}>{chart.config.yAxis.join(', ')} vs {chart.config.xAxis}</p>
                            <p className={`text-[10px] ${textMuted}`}>{typeLabel}</p>
                          </div>
                          <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full bg-emerald-500/20 text-emerald-400 whitespace-nowrap">✓ Saved</span>
                        </div>
                        <p className={`text-[10px] leading-relaxed mb-2 line-clamp-1 ${textMuted}`}>💡 {insight}</p>
                        <div className="flex items-center justify-between">
                          <div className="flex gap-1">
                            <button onClick={() => handleMoveChart(idx, -1)} disabled={idx === 0} className={`p-1 rounded text-xs transition-colors ${dark ? 'text-slate-500 hover:text-slate-300 disabled:opacity-30' : 'text-slate-400 hover:text-slate-600 disabled:opacity-30'}`}>↑</button>
                            <button onClick={() => handleMoveChart(idx, 1)} disabled={idx === savedCharts.length - 1} className={`p-1 rounded text-xs transition-colors ${dark ? 'text-slate-500 hover:text-slate-300 disabled:opacity-30' : 'text-slate-400 hover:text-slate-600 disabled:opacity-30'}`}>↓</button>
                          </div>
                          <button onClick={() => handleRemoveChart(chart.id)} className={`text-xs px-2 py-1 rounded transition-colors ${dark ? 'text-slate-500 hover:text-red-400 hover:bg-red-500/10' : 'text-slate-400 hover:text-red-500 hover:bg-red-50'}`}>🗑️</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </aside>
        )}
      </div>
    </div>
  );
}
