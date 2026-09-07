const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { createClient } = require('@supabase/supabase-js');

const workspaceDir = 'c:\\Users\\User\\Desktop\\Ambacar_Proformas';

// Read Supabase credentials dynamically from .env
let supabaseUrl = "https://mblsjoreokpsjtdzjrae.supabase.co";
let supabaseKey = "";

const envPath = path.join(workspaceDir, '.env');
if (fs.existsSync(envPath)) {
  const envContent = fs.readFileSync(envPath, 'utf8');
  for (const line of envContent.split('\n')) {
    const parts = line.split('=');
    if (parts.length >= 2) {
      const k = parts[0].trim();
      const v = parts.slice(1).join('=').trim();
      if (k === 'VITE_SUPABASE_URL') supabaseUrl = v;
      if (k === 'VITE_SUPABASE_ANON_KEY') supabaseKey = v;
    }
  }
}

const supabase = supabaseKey ? createClient(supabaseUrl, supabaseKey) : null;

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

// Model Definitions & mappings to sheets in the 5 Excels
const planDefinitions = [
  // 1. Wingle 7 (250K)
  {
    file: "PLAN MANT W7 - 5000 hasta 250000KM.XLSX",
    sheet: "DETALLADO MANT 4X4",
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
    marca: "GWM",
    modelo: "WINGLE 2.8",
    motor: "2.8 DIESEL",
    traccion: "4X4 / 4X2",
    nombre_completo: "WINGLE 2.8 DIESEL",
    codigo_plan: "W28-DIE-120K",
    nombre_plan: "Plan de Mantenimiento Wingle 2.8 Diesel (5.000 - 120.000 KM)"
  },

  // 2. POER Diesel (200K)
  {
    file: "PLAN MANT POER DIE 4X4 Y 4X2 desde 5000km hasta 200.000 KM.xlsx",
    sheet: "POER AC 2.0 CD 4X4 TM DIESEL",
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
    marca: "GWM",
    modelo: "POER",
    motor: "2.0 DIESEL",
    traccion: "4X2",
    nombre_completo: "GWM POER DIESEL 4X2",
    codigo_plan: "POER-DIE-4X2-200K",
    nombre_plan: "Plan de Mantenimiento GWM POER 2.0 CD Diesel 4x2 (5.000 - 200.000 KM)"
  },

  // 3. POER Policia & Gasolina
  {
    file: "PLAN MANT POER POER POLICIA.xlsx",
    sheet: "Hoja1",
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
    marca: "GWM",
    modelo: "POER",
    motor: "2.0 DIESEL",
    traccion: "4X4 POLICIA",
    nombre_completo: "GWM POER DIESEL PATRULLEROS POLICIA",
    codigo_plan: "POER-POLICIA-120K",
    nombre_plan: "Plan de Mantenimiento POER Patrulleros Policía (5.000 - 120.000 KM)"
  },

  // 4. TANK 300 y TANK 500
  {
    file: "PLAN MANT TANK 300 500 HASTA 120.000 KM.xlsx",
    sheet: "TANK 300",
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
    marca: "TANK",
    modelo: "TANK 500",
    motor: "3.0 V6 TURBO",
    traccion: "4X4",
    nombre_completo: "TANK 500 4X4",
    codigo_plan: "TANK500-120K",
    nombre_plan: "Plan de Mantenimiento TANK 500 (5.000 - 120.000 KM)"
  },

  // 5. KYC F3
  {
    file: "PLAN MANT KYC F3 - ACT 2024 - CIAUTO ref. 15-5-2025.XLSX",
    sheet: "DETALLADO MANT 4X4",
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
    marca: "KYC",
    modelo: "KYC F3",
    motor: "GASOLINA",
    traccion: "4X2",
    nombre_completo: "KYC F3 4X2",
    codigo_plan: "KYC-F3-4X2-120K",
    nombre_plan: "Plan de Mantenimiento KYC F3 4X2 (5.000 - 120.000 KM)"
  }
];

// Corrective Catalogs
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

async function generateCompactSql() {
  console.log("==================================================================");
  console.log("GENERANDO SQL COMPACTO OPTIMIZADO PARA EL EDITOR SQL DE SUPABASE");
  console.log("==================================================================");

  let sql = [];
  sql.push(`-- Limpieza previa`);
  sql.push(`TRUNCATE TABLE public.plan_items_detalle CASCADE;`);
  sql.push(`TRUNCATE TABLE public.plan_costos_km CASCADE;`);
  sql.push(`TRUNCATE TABLE public.planes_mantenimiento CASCADE;`);
  sql.push(`TRUNCATE TABLE public.catalogo_correctivo CASCADE;`);
  sql.push(`TRUNCATE TABLE public.modelos_vehiculo CASCADE;\n`);

  // 1. Modelos
  const modelosRows = [];
  const modelosSet = new Set();
  for (const def of planDefinitions) {
    if (!modelosSet.has(def.nombre_completo)) {
      modelosSet.add(def.nombre_completo);
      modelosRows.push(`('${def.marca}', '${def.modelo}', '${def.motor}', '${def.traccion}', '${def.nombre_completo}')`);
    }
  }
  sql.push(`INSERT INTO public.modelos_vehiculo (marca, modelo, motor, traccion, nombre_completo) VALUES\n` + modelosRows.join(',\n') + `\nON CONFLICT (nombre_completo) DO NOTHING;\n`);

  // 2. Planes y Costos por KM
  const planesRows = [];
  const costosKmRows = [];
  const itemsRows = [];

  for (const def of planDefinitions) {
    const filePath = path.join(workspaceDir, def.file);
    if (!fs.existsSync(filePath)) continue;

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[def.sheet];
    if (!ws) continue;

    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Header row with KMs
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

    const kmMin = Math.min(...kmCols.map(x => x.km));
    const kmMax = Math.max(...kmCols.map(x => x.km));

    planesRows.push(`((SELECT id FROM public.modelos_vehiculo WHERE nombre_completo = '${def.nombre_completo}'), '${def.codigo_plan}', '${def.nombre_plan}', '${def.file}', '${def.sheet}', ${kmMin}, ${kmMax}, 5000)`);

    // Totals rows
    let filaRep = null, filaLub = null, filaMo = null, filaTot = null;
    const itemsList = [];
    let currentCategoria = 'REPUESTOS';

    for (let r = 0; r < grid.length; r++) {
      const c0 = String(grid[r]?.[0] || '').trim().toUpperCase();
      if (c0.includes('TOTAL LUBRICANTE') || c0 === 'TOTAL LUBRICANTES') { filaLub = r; currentCategoria = 'REPUESTOS'; continue; }
      if (c0.includes('TOTAL REPUESTO') || c0 === 'TOTAL REPUESTOS') { filaRep = r; currentCategoria = 'MANO DE OBRA'; continue; }
      if (c0.includes('TOTAL MANO DE OBRA') || c0.includes('TOTAL M/O')) { filaMo = r; continue; }
      if (c0.includes('TOTAL COSTO') || c0.includes('TOTAL MANTENIMIENTO') || c0 === 'TOTAL') { filaTot = r; continue; }

      if (r > kmRowIndex && c0 && !c0.includes('PLAN DE') && !c0.includes('MARCA') && !c0.includes('MODELO')) {
        const hasNumbers = kmCols.some(k => grid[r]?.[k.col] !== '' && grid[r]?.[k.col] !== undefined);
        if (hasNumbers) {
          let cat = currentCategoria;
          if (c0.includes('ACEITE') || c0.includes('LUBRIC') || c0.includes('LIQUIDO') || c0.includes('REFRIGERANTE')) { cat = 'LUBRICANTES'; }
          itemsList.push({ categoria: cat, descripcion: String(grid[r][0]).trim(), row: r });
        }
      }
    }

    for (const { col, km } of kmCols) {
      const rep = filaRep !== null ? parseNumber(grid[filaRep]?.[col]) : 0.00;
      const lub = filaLub !== null ? parseNumber(grid[filaLub]?.[col]) : 0.00;
      const mo = filaMo !== null ? parseNumber(grid[filaMo]?.[col]) : 0.00;
      const tot = filaTot !== null ? parseNumber(grid[filaTot]?.[col]) : (rep + lub + mo);

      costosKmRows.push(`((SELECT id FROM public.planes_mantenimiento WHERE codigo_plan = '${def.codigo_plan}'), ${km}, ${rep}, ${lub}, ${mo}, ${tot})`);

      for (const item of itemsList) {
        const val = grid[item.row]?.[col];
        const num = parseNumber(val);
        const aplica = (val === 'X' || val === 'x' || num > 0);
        if (aplica) {
          const descEsc = item.descripcion.replace(/'/g, "''");
          itemsRows.push(`((SELECT id FROM public.planes_mantenimiento WHERE codigo_plan = '${def.codigo_plan}'), '${item.categoria}', '${descEsc}', ${km}, ${num}, ${aplica})`);
        }
      }
    }
  }

  // Insert Planes
  sql.push(`INSERT INTO public.planes_mantenimiento (vehiculo_id, codigo_plan, nombre, archivo_origen, hoja_origen, km_inicio, km_fin, intervalo_km) VALUES\n` + planesRows.join(',\n') + `\nON CONFLICT (codigo_plan) DO NOTHING;\n`);

  // Insert Costos por KM (batched by 200)
  for (let i = 0; i < costosKmRows.length; i += 200) {
    const chunk = costosKmRows.slice(i, i + 200);
    sql.push(`INSERT INTO public.plan_costos_km (plan_id, km, total_repuestos, total_lubricantes, total_mano_obra, total_km) VALUES\n` + chunk.join(',\n') + `\nON CONFLICT (plan_id, km) DO UPDATE SET total_repuestos = EXCLUDED.total_repuestos, total_lubricantes = EXCLUDED.total_lubricantes, total_mano_obra = EXCLUDED.total_mano_obra, total_km = EXCLUDED.total_km;\n`);
  }

  // 3. Catálogos Correctivos
  const corrRows = [];
  for (const corr of correctiveDefinitions) {
    const filePath = path.join(workspaceDir, corr.file);
    if (!fs.existsSync(filePath)) continue;

    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[corr.sheet];
    if (!ws) continue;

    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    let startRow = 2;
    for (let r = 0; r < Math.min(10, grid.length); r++) {
      const row = grid[r] || [];
      if (row.some(c => String(c).toUpperCase().includes('DESCRIPCION') || String(c).toUpperCase().includes('DESCRIPCIÓN'))) {
        startRow = r + 1;
        break;
      }
    }

    for (let r = startRow; r < grid.length; r++) {
      const row = grid[r] || [];
      const itemNum = parseInt(row[0], 10);
      const desc = String(row[1] || '').trim();
      
      if (!desc || desc.toUpperCase().includes('TOTAL') || desc.toUpperCase().includes('SUBTOTAL') || desc.toUpperCase().includes('NOTA')) continue;

      const rep = parseNumber(row[2]);
      const mo = parseNumber(row[3]);
      const lub = parseNumber(row[4]);
      const sub = parseNumber(row[5]) || (rep + mo + lub);
      const iva = parseNumber(row[6]) || (Math.round(sub * 0.15 * 100) / 100);
      const tot = parseNumber(row[7]) || (sub + iva);

      const descEsc = desc.replace(/'/g, "''");
      corrRows.push(`('${corr.modelo_referencia}', '${corr.tipo_catalogo}', ${isNaN(itemNum) ? 'NULL' : itemNum}, '${descEsc}', ${rep}, ${mo}, ${lub}, ${sub}, ${iva}, ${tot})`);
    }
  }

  if (corrRows.length > 0) {
    sql.push(`INSERT INTO public.catalogo_correctivo (modelo_referencia, tipo_catalogo, item_numero, descripcion, costo_repuestos, costo_mano_obra, costo_lubricantes, subtotal, iva, total) VALUES\n` + corrRows.join(',\n') + `;\n`);
  }

  // Write Setup Compact (Schema + Seed without individual 6k items, just models, plans, and KM totals)
  const schemaPath = path.join(workspaceDir, 'supabase', 'schema.sql');
  const schemaContent = fs.readFileSync(schemaPath, 'utf8');

  const setupCompactPath = path.join(workspaceDir, 'supabase', 'setup_compact.sql');
  const finalContent = schemaContent + '\n\n' + sql.join('\n');
  fs.writeFileSync(setupCompactPath, finalContent, 'utf-8');

  const stats = fs.statSync(setupCompactPath);
  console.log(`✅ Archivo compacto generado: ${setupCompactPath}`);
  console.log(`   Tamaño optimizado: ${(stats.size / 1024).toFixed(2)} KB (¡Completamente admitido por el SQL Editor de Supabase!)`);
}

generateCompactSql().catch(console.error);
