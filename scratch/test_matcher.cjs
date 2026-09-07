const { PLANES_MANTENIMIENTO_DATA } = require('../src/data/planesData.js');

const DEFAULT_MODELOS = PLANES_MANTENIMIENTO_DATA.map(p => ({
  id: p.id,
  marca: p.marca,
  modelo: p.modelo,
  motor: p.motor,
  traccion: p.traccion,
  nombre_completo: p.nombre_completo,
  planes_mantenimiento: [{
    id: p.id,
    codigo_plan: p.codigo_plan,
    nombre: p.nombre_plan,
    km_inicio: p.km_inicio,
    km_fin: p.km_fin,
    intervalo_km: 5000
  }]
}));

async function getModelosVehiculo() {
  return { data: DEFAULT_MODELOS, error: null, source: 'local' };
}

async function buscarPlanPorModelo(modeloTexto = '', observacionTexto = '') {
  const combined = `${modeloTexto || ''} ${observacionTexto || ''}`.toUpperCase().trim();
  const normalized = (modeloTexto || '').toUpperCase().trim();
  const { data: modelos } = await getModelosVehiculo();
  
  if (!modelos || modelos.length === 0) {
    return { plan: DEFAULT_MODELOS[0].planes_mantenimiento[0], modelo: DEFAULT_MODELOS[0] };
  }

  let bestMatch = null;
  let highestScore = -1;

  for (const m of modelos) {
    let score = 0;
    const mName = m.nombre_completo.toUpperCase();
    const mModelo = m.modelo.toUpperCase();
    const mTraccion = (m.traccion || '').toUpperCase();
    const mMotor = (m.motor || '').toUpperCase();

    // Exact full match
    if (normalized === mName || normalized === mModelo) score += 200;
    if (mName.includes(normalized) && normalized.length > 3) score += 60;
    if (normalized.includes(mName)) score += 60;

    // Specific Model families
    const hasTank300 = combined.includes('TANK 300') || combined.includes('TANK300') || (combined.includes('TANK') && combined.includes('300'));
    const hasTank500 = combined.includes('TANK 500') || combined.includes('TANK500') || (combined.includes('TANK') && combined.includes('500'));
    const hasWingle28 = combined.includes('2.8') || combined.includes('2,8');
    const hasWingle7 = combined.includes('W7') || combined.includes('WINGLE 7') || combined.includes('WINGLE7') || (combined.includes('WINGLE') && !hasWingle28);
    const hasPoer = combined.includes('POER');
    const hasKyc = combined.includes('KYC') || combined.includes('F3');

    if (mModelo.includes('TANK 300') && hasTank300) score += 180;
    if (mModelo.includes('TANK 500') && hasTank500) score += 180;
    if (mModelo.includes('WINGLE 2.8') && hasWingle28) score += 180;
    if (mModelo.includes('WINGLE 7') && hasWingle7) score += 120;
    if (mModelo.includes('POER') && hasPoer) score += 120;
    if (mModelo.includes('KYC') && hasKyc) score += 150;

    // Traction matching (4x4 vs 4x2)
    const has4x2 = combined.includes('4X2') || combined.includes('4 X 2') || combined.includes('2WD') || combined.includes('SIMPLE');
    const has4x4 = combined.includes('4X4') || combined.includes('4 X 4') || combined.includes('4WD') || combined.includes('DOBLE TRACCION');

    if (has4x2 && (mTraccion.includes('4X2') || mName.includes('4X2'))) score += 60;
    if (has4x4 && (mTraccion.includes('4X4') || mName.includes('4X4'))) score += 60;

    // Special vehicle variants
    if ((combined.includes('POLIC') || combined.includes('PATRULL')) && mName.includes('POLIC')) score += 80;
    if (combined.includes('GASOLINA') && (mMotor.includes('GASOLINA') || mName.includes('GASOLINA'))) score += 50;
    if (combined.includes('DIESEL') && (mMotor.includes('DIESEL') || mName.includes('DIESEL'))) score += 20;

    console.log(`${m.nombre_completo.padEnd(35)} -> score: ${score}`);
    if (score > highestScore && score > 0) {
      highestScore = score;
      bestMatch = m;
    }
  }

  return { plan: bestMatch?.planes_mantenimiento[0], modelo: bestMatch };
}

(async () => {
  console.log("--- TEST 1: modeloTexto = 'TANK 300' ---");
  const res1 = await buscarPlanPorModelo('TANK 300', '');
  console.log('WINNER:', res1.modelo.nombre_completo);

  console.log("\n--- TEST 2: modeloTexto = 'TANK 300 4X4' ---");
  const res2 = await buscarPlanPorModelo('TANK 300 4X4', '');
  console.log('WINNER:', res2.modelo.nombre_completo);

  console.log("\n--- TEST 3: grid search simulation ---");
  // If the cell in Excel has 'TANK 300'
  const mockGrid = [
    ["AMBACAR S.A.", "", ""],
    ["Razón Social:", "MINISTERIO DEL INTERIOR"],
    ["Vigencia de la oferta:", 90, "días."],
    ["Modelo de vehículo:", "TANK 300"]
  ];
  
  function buscarValor(grid, etiquetas) {
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      for (let c = 0; c < row.length; c++) {
        const cellText = (row[c] || "").toString().trim().toLowerCase();
        for (const etiq of etiquetas) {
          const etiqNorm = etiq.toLowerCase().trim();
          if (cellText === etiqNorm || cellText.startsWith(etiqNorm) || cellText.includes(etiqNorm)) {
            for (let valCol = c + 1; valCol < row.length; valCol++) {
              const v = (row[valCol] || "").toString().trim();
              if (v && !["nan", "none", "", "null", "undefined", "-", ":"].includes(v.toLowerCase())) {
                return v;
              }
            }
          }
        }
      }
    }
    return null;
  }

  const found = buscarValor(mockGrid, ["Modelo de vehículo", "Modelo de vehiculo", "Modelo"]);
  console.log("Found in mockGrid:", found);
})();
