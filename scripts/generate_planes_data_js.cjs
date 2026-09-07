const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const workspaceDir = 'c:\\Users\\User\\Desktop\\Ambacar_Proformas';

function kmAInt(val) {
  if (val === null || val === undefined) return null;
  const str = String(val).toUpperCase().replace(/KM/g, '').replace(/\./g, '').replace(/,/g, '').trim();
  const parsed = parseInt(str, 10);
  return isNaN(parsed) ? null : parsed;
}

function parseNumber(val) {
  if (val === null || val === undefined || val === '' || val === 'X' || val === 'x') return 0.00;
  if (typeof val === 'number') return Math.round(val * 100) / 100;
  const clean = String(val).replace(/,/g, '').replace(/\$/g, '').trim();
  const parsed = parseFloat(clean);
  return isNaN(parsed) ? 0.00 : Math.round(parsed * 100) / 100;
}

const planDefinitions = [
  {
    file: "PLAN MANT W7 - 5000 hasta 250000KM.XLSX",
    sheet: "DETALLADO MANT 4X4",
    id: "w7-4x4-250k",
    marca: "GWM",
    modelo: "WINGLE 7",
    motor: "2.0 DIESEL",
    traccion: "4X4",
    nombre_completo: "WINGLE 7 DIESEL 4X4",
    codigo_plan: "W7-DIE-4X4-250K",
    nombre_plan: "Plan de Mantenimiento Wingle 7 Diesel 4x4 (5.000 - 250.000 KM)"
  },
  {
    file: "PLAN MANT W7 - 5000 hasta 250000KM.XLSX",
    sheet: "DETALLADO MANT 4X2.",
    id: "w7-4x2-250k",
    marca: "GWM",
    modelo: "WINGLE 7",
    motor: "2.0 DIESEL",
    traccion: "4X2",
    nombre_completo: "WINGLE 7 DIESEL 4X2",
    codigo_plan: "W7-DIE-4X2-250K",
    nombre_plan: "Plan de Mantenimiento Wingle 7 Diesel 4x2 (5.000 - 250.000 KM)"
  },
  {
    file: "PLAN MANT W7 - 5000 hasta 250000KM.XLSX",
    sheet: "PLAN MANTENIMEINTO WINGLE 2,8",
    id: "w28-120k",
    marca: "GWM",
    modelo: "WINGLE 2.8",
    motor: "2.8 DIESEL",
    traccion: "4X4 / 4X2",
    nombre_completo: "WINGLE 2.8 DIESEL",
    codigo_plan: "W28-DIE-120K",
    nombre_plan: "Plan de Mantenimiento Wingle 2.8 Diesel (5.000 - 120.000 KM)"
  },
  {
    file: "PLAN MANT POER DIE 4X4 Y 4X2 desde 5000km hasta 200.000 KM.xlsx",
    sheet: "POER AC 2.0 CD 4X4 TM DIESEL",
    id: "poer-4x4-200k",
    marca: "GWM",
    modelo: "POER",
    motor: "2.0 DIESEL",
    traccion: "4X4",
    nombre_completo: "GWM POER DIESEL 4X4",
    codigo_plan: "POER-DIE-4X4-200K",
    nombre_plan: "Plan de Mantenimiento GWM POER 2.0 CD Diesel 4x4 (5.000 - 200.000 KM)"
  },
  {
    file: "PLAN MANT POER DIE 4X4 Y 4X2 desde 5000km hasta 200.000 KM.xlsx",
    sheet: "POER AC 2.0 CD 4X2 TM DIESE",
    id: "poer-4x2-200k",
    marca: "GWM",
    modelo: "POER",
    motor: "2.0 DIESEL",
    traccion: "4X2",
    nombre_completo: "GWM POER DIESEL 4X2",
    codigo_plan: "POER-DIE-4X2-200K",
    nombre_plan: "Plan de Mantenimiento GWM POER 2.0 CD Diesel 4x2 (5.000 - 200.000 KM)"
  },
  {
    file: "PLAN MANT POER POER POLICIA.xlsx",
    sheet: "Hoja1",
    id: "poer-gas-120k",
    marca: "GWM",
    modelo: "POER",
    motor: "GASOLINA",
    traccion: "4X4",
    nombre_completo: "GWM POER GASOLINA",
    codigo_plan: "POER-GAS-120K",
    nombre_plan: "Plan de Mantenimiento GWM POER Gasolina (5.000 - 120.000 KM)"
  },
  {
    file: "PLAN MANT POER POER POLICIA.xlsx",
    sheet: "POER AC 2.0 CD 4X4 TM DIESEL",
    id: "poer-policia-120k",
    marca: "GWM",
    modelo: "POER",
    motor: "2.0 DIESEL",
    traccion: "4X4 POLICIA",
    nombre_completo: "GWM POER DIESEL PATRULLEROS POLICIA",
    codigo_plan: "POER-POLICIA-120K",
    nombre_plan: "Plan de Mantenimiento POER Patrulleros Policía (5.000 - 120.000 KM)"
  },
  {
    file: "PLAN MANT TANK 300 500 HASTA 120.000 KM.xlsx",
    sheet: "TANK 300",
    id: "tank-300-120k",
    marca: "TANK",
    modelo: "TANK 300",
    motor: "2.0 TURBO",
    traccion: "4X4",
    nombre_completo: "TANK 300 4X4",
    codigo_plan: "TANK300-120K",
    nombre_plan: "Plan de Mantenimiento TANK 300 (5.000 - 120.000 KM)"
  },
  {
    file: "PLAN MANT TANK 300 500 HASTA 120.000 KM.xlsx",
    sheet: "TANK 500",
    id: "tank-500-120k",
    marca: "TANK",
    modelo: "TANK 500",
    motor: "3.0 V6 TURBO",
    traccion: "4X4",
    nombre_completo: "TANK 500 4X4",
    codigo_plan: "TANK500-120K",
    nombre_plan: "Plan de Mantenimiento TANK 500 (5.000 - 120.000 KM)"
  },
  {
    file: "PLAN MANT KYC F3 - ACT 2024 - CIAUTO ref. 15-5-2025.XLSX",
    sheet: "DETALLADO MANT 4X4",
    id: "kyc-f3-4x4-120k",
    marca: "KYC",
    modelo: "KYC F3",
    motor: "GASOLINA",
    traccion: "4X4",
    nombre_completo: "KYC F3 4X4",
    codigo_plan: "KYC-F3-4X4-120K",
    nombre_plan: "Plan de Mantenimiento KYC F3 4X4 (5.000 - 120.000 KM)"
  },
  {
    file: "PLAN MANT KYC F3 - ACT 2024 - CIAUTO ref. 15-5-2025.XLSX",
    sheet: "DETALLADO MANT 4X2",
    id: "kyc-f3-4x2-120k",
    marca: "KYC",
    modelo: "KYC F3",
    motor: "GASOLINA",
    traccion: "4X2",
    nombre_completo: "KYC F3 4X2",
    codigo_plan: "KYC-F3-4X2-120K",
    nombre_plan: "Plan de Mantenimiento KYC F3 4X2 (5.000 - 120.000 KM)"
  }
];

const correctiveDefinitions = [
  {
    file: "PLAN MANT W7 - 5000 hasta 250000KM.XLSX",
    sheet: "REF. MANT. CORRECTIVO",
    modelo_referencia: "WINGLE 7 DIESEL",
    tipo_catalogo: "ESTANDAR"
  },
  {
    file: "PLAN MANT W7 - 5000 hasta 250000KM.XLSX",
    sheet: "REF. MANT. CORRECTIVO COMPLETO",
    modelo_referencia: "WINGLE 7 DIESEL",
    tipo_catalogo: "COMPLETO"
  },
  {
    file: "PLAN MANT POER DIE 4X4 Y 4X2 desde 5000km hasta 200.000 KM.xlsx",
    sheet: "REF MANT CORRECTIVO",
    modelo_referencia: "POER DIESEL",
    tipo_catalogo: "ESTANDAR"
  },
  {
    file: "PLAN MANT TANK 300 500 HASTA 120.000 KM.xlsx",
    sheet: "REF MANT CORRECTIVO",
    modelo_referencia: "TANK 300 / 500",
    tipo_catalogo: "ESTANDAR"
  },
  {
    file: "PLAN MANT KYC F3 - ACT 2024 - CIAUTO ref. 15-5-2025.XLSX",
    sheet: "CORRECTIVO",
    modelo_referencia: "KYC F3",
    tipo_catalogo: "ESTANDAR"
  }
];

const planesOutput = [];

for (const def of planDefinitions) {
  const filePath = path.join(workspaceDir, def.file);
  if (!fs.existsSync(filePath)) continue;

  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[def.sheet];
  if (!ws) continue;

  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let kmRowIndex = null;
  let kmCols = [];

  for (let r = 0; r < Math.min(10, grid.length); r++) {
    const row = grid[r] || [];
    const foundKms = [];
    for (let c = 1; c < row.length; c++) {
      const val = kmAInt(row[c]);
      if (val && val >= 5000) {
        foundKms.push({ col: c, km: val });
      }
    }
    if (foundKms.length >= 3) {
      kmRowIndex = r;
      kmCols = foundKms;
      break;
    }
  }

  if (!kmCols.length) continue;

  let filaRep = null, filaLub = null, filaMo = null, filaTot = null;

  for (let r = 0; r < grid.length; r++) {
    const c0 = String(grid[r]?.[0] || '').trim().toUpperCase();
    if (c0.includes('TOTAL LUBRICANTE') || c0 === 'TOTAL LUBRICANTES') { filaLub = r; continue; }
    if (c0.includes('TOTAL REPUESTO') || c0 === 'TOTAL REPUESTOS') { filaRep = r; continue; }
    if (c0.includes('TOTAL MANO DE OBRA') || c0.includes('TOTAL M/O')) { filaMo = r; continue; }
    if (c0.includes('TOTAL COSTO') || c0.includes('TOTAL MANTENIMIENTO') || c0 === 'TOTAL') { filaTot = r; continue; }
  }

  const costosPorKm = {};

  for (const { col, km } of kmCols) {
    const rep = filaRep !== null ? parseNumber(grid[filaRep]?.[col]) : 0.00;
    const lub = filaLub !== null ? parseNumber(grid[filaLub]?.[col]) : 0.00;
    const mo = filaMo !== null ? parseNumber(grid[filaMo]?.[col]) : 0.00;
    const tot = filaTot !== null ? parseNumber(grid[filaTot]?.[col]) : (rep + lub + mo);

    costosPorKm[km] = {
      km,
      total_repuestos: rep,
      total_lubricantes: lub,
      total_mano_obra: mo,
      total_km: tot
    };
  }

  planesOutput.push({
    id: def.id,
    marca: def.marca,
    modelo: def.modelo,
    motor: def.motor,
    traccion: def.traccion,
    nombre_completo: def.nombre_completo,
    codigo_plan: def.codigo_plan,
    nombre_plan: def.nombre_plan,
    km_inicio: Math.min(...kmCols.map(x => x.km)),
    km_fin: Math.max(...kmCols.map(x => x.km)),
    costosPorKm
  });
}

// Ensure dir
const dataDir = path.join(workspaceDir, 'src', 'data');
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}

const fileContent = `// Pre-extracted offline/cloud dataset from official Excel maintenance plans (AMBACAR)
export const PLANES_MANTENIMIENTO_DATA = ${JSON.stringify(planesOutput, null, 2)};
`;

fs.writeFileSync(path.join(dataDir, 'planesData.js'), fileContent, 'utf-8');
console.log(`✅ Archivo generado: src/data/planesData.js con ${planesOutput.length} modelos de vehículos y todos sus costos.`);
