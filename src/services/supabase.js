import { createClient } from '@supabase/supabase-js';
import { PLANES_MANTENIMIENTO_DATA } from '../data/planesData';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || "https://mblsjoreokpsjtdzjrae.supabase.co";
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || "";

export const supabase = createClient(supabaseUrl, supabaseAnonKey || "dummy-anon-key");

export const DEFAULT_MODELOS = PLANES_MANTENIMIENTO_DATA.map(p => ({
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

/**
 * Fetches all vehicle models.
 * Uses Supabase when online/configured, with automatic fallback to bundled Excel dataset.
 */
export async function getModelosVehiculo() {
  if (supabaseAnonKey && supabaseAnonKey !== "dummy-anon-key") {
    try {
      const { data, error } = await supabase
        .from('modelos_vehiculo')
        .select('*, planes_mantenimiento(*)')
        .order('marca', { ascending: true })
        .order('modelo', { ascending: true });
        
      if (!error && data && data.length > 0) {
        return { data, error: null, source: 'supabase' };
      }
    } catch (err) {
      console.warn("Supabase fetch failed, using pre-loaded dataset:", err);
    }
  }

  // Instant fallback to bundled dataset from the 5 Excels
  return { data: DEFAULT_MODELOS, error: null, source: 'local' };
}

/**
 * Robust vehicle model matcher for client quotation text.
 */
export async function buscarPlanPorModelo(modeloTexto = "", observacionTexto = "") {
  const combined = `${modeloTexto || ''} ${observacionTexto || ''}`.toUpperCase().trim();
  const normalized = (modeloTexto || "").toUpperCase().trim();
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
    const hasTank300 = combined.includes("TANK 300") || combined.includes("TANK300") || (combined.includes("TANK") && combined.includes("300"));
    const hasTank500 = combined.includes("TANK 500") || combined.includes("TANK500") || (combined.includes("TANK") && combined.includes("500"));
    const hasWingle28 = combined.includes("2.8") || combined.includes("2,8");
    const hasWingle7 = combined.includes("W7") || combined.includes("WINGLE 7") || combined.includes("WINGLE7") || (combined.includes("WINGLE") && !hasWingle28);
    const hasPoer = combined.includes("POER");
    const hasKyc = combined.includes("KYC") || combined.includes("F3");

    if (mModelo.includes("TANK 300") && hasTank300) score += 180;
    if (mModelo.includes("TANK 500") && hasTank500) score += 180;
    if (mModelo.includes("WINGLE 2.8") && hasWingle28) score += 180;
    if (mModelo.includes("WINGLE 7") && hasWingle7) score += 120;
    if (mModelo.includes("POER") && hasPoer) score += 120;
    if (mModelo.includes("KYC") && hasKyc) score += 150;

    // Traction matching (4x4 vs 4x2)
    const has4x2 = combined.includes("4X2") || combined.includes("4 X 2") || combined.includes("2WD") || combined.includes("SIMPLE");
    const has4x4 = combined.includes("4X4") || combined.includes("4 X 4") || combined.includes("4WD") || combined.includes("DOBLE TRACCION");

    if (has4x2 && (mTraccion.includes("4X2") || mName.includes("4X2"))) score += 60;
    if (has4x4 && (mTraccion.includes("4X4") || mName.includes("4X4"))) score += 60;

    // Special vehicle variants
    if ((combined.includes("POLIC") || combined.includes("PATRULL")) && mName.includes("POLIC")) score += 80;
    if (combined.includes("GASOLINA") && (mMotor.includes("GASOLINA") || mName.includes("GASOLINA"))) score += 50;
    if (combined.includes("DIESEL") && (mMotor.includes("DIESEL") || mName.includes("DIESEL"))) score += 20;

    if (score > highestScore && score > 0) {
      highestScore = score;
      bestMatch = m;
    }
  }

  // Fallback default to first model if no specific match
  if (!bestMatch && modelos.length > 0) {
    bestMatch = modelos[0];
  }

  if (bestMatch && bestMatch.planes_mantenimiento && bestMatch.planes_mantenimiento.length > 0) {
    return { plan: bestMatch.planes_mantenimiento[0], modelo: bestMatch };
  }

  return { plan: DEFAULT_MODELOS[0].planes_mantenimiento[0], modelo: DEFAULT_MODELOS[0] };
}

function roundMoney(val) {
  return Math.round((Number(val || 0) + 0.00001) * 100) / 100;
}

/**
 * Calculates maintenance costs for a given plan and KM range or individual vehicle maintenance list.
 */
export async function getCostosPorKm(planId, kmDesde, kmHasta, cantidadVehiculos = 1, mantenimientosEspecificos = null) {
  const kDesde = parseInt(kmDesde, 10) || 0;
  const kHasta = parseInt(kmHasta, 10) || 250000;
  const nVehiculos = parseInt(cantidadVehiculos, 10) || 1;

  let localPlan = PLANES_MANTENIMIENTO_DATA.find(p => p.id === planId || p.codigo_plan === planId);
  
  // If not found directly, search through DEFAULT_MODELOS and Supabase models
  if (!localPlan) {
    const { data: modelos } = await getModelosVehiculo();
    const matchedMod = modelos?.find(m => m.id === planId || m.planes_mantenimiento?.some(pm => pm.id === planId || pm.codigo_plan === planId));
    if (matchedMod) {
      localPlan = PLANES_MANTENIMIENTO_DATA.find(p => 
        p.nombre_completo?.toUpperCase() === matchedMod.nombre_completo?.toUpperCase() ||
        p.modelo?.toUpperCase() === matchedMod.modelo?.toUpperCase()
      );
    }
  }

  if (!localPlan) {
    localPlan = PLANES_MANTENIMIENTO_DATA[0];
  }

  // 1. If specific individual vehicle maintenance list is given (e.g. from Observación)
  if (mantenimientosEspecificos && Array.isArray(mantenimientosEspecificos) && mantenimientosEspecificos.length > 0) {
    let totalRepuestos = 0;
    let totalLubricantes = 0;
    let totalMo = 0;

    const tablaDetalle = mantenimientosEspecificos.map((item, index) => {
      const km = item.km;
      const dataKm = localPlan.costosPorKm[km] || { total_repuestos: 0, total_lubricantes: 0, total_mano_obra: 0, total_km: 0 };
      const rep = parseFloat(dataKm.total_repuestos) || 0;
      const lub = parseFloat(dataKm.total_lubricantes) || 0;
      const mo = parseFloat(dataKm.total_mano_obra) || 0;
      const tot = parseFloat(dataKm.total_km) || (rep + lub + mo);

      totalRepuestos += rep;
      totalLubricantes += lub;
      totalMo += mo;

      return {
        itemNum: index + 1,
        placa: item.placa || `Vehículo ${index + 1}`,
        km,
        repuestos: rep,
        lubricantes: lub,
        mano_obra: mo,
        total: tot
      };
    });

    const totalPreventivo = roundMoney(totalRepuestos + totalLubricantes + totalMo);
    totalRepuestos = roundMoney(totalRepuestos);
    totalLubricantes = roundMoney(totalLubricantes);
    totalMo = roundMoney(totalMo);

    return {
      error: false,
      modo: 'individual',
      mantenimientosEspecificos,
      tablaDetalle,
      nVehiculos: mantenimientosEspecificos.length,
      totalRepuestos1v: totalRepuestos,
      totalLubricantes1v: totalLubricantes,
      totalMo1v: totalMo,
      totalPreventivo1v: totalPreventivo,
      totalRepuestos,
      totalLubricantes,
      totalMo,
      totalPreventivo
    };
  }

  // 2. Standard range calculation
  const allKms = Object.keys(localPlan.costosPorKm).map(Number).sort((a, b) => a - b);
  const matchingKms = allKms.filter(k => k >= kDesde && k <= kHasta);

  if (matchingKms.length === 0) {
    return {
      error: true,
      kmsDisponibles: allKms,
      message: `No hay mantenimientos en el rango [${kDesde} - ${kHasta} KM]. Kilómetros disponibles: ${allKms.join(', ')}`
    };
  }

  let totalRepuestos1v = 0;
  let totalLubricantes1v = 0;
  let totalMo1v = 0;

  const tablaDetalle = matchingKms.map(km => {
    const item = localPlan.costosPorKm[km];
    const rep = item.total_repuestos;
    const lub = item.total_lubricantes;
    const mo = item.total_mano_obra;
    const tot = item.total_km || (rep + lub + mo);

    totalRepuestos1v += rep;
    totalLubricantes1v += lub;
    totalMo1v += mo;

    return { km, repuestos: rep, lubricantes: lub, mano_obra: mo, total: tot };
  });

  totalRepuestos1v = roundMoney(totalRepuestos1v);
  totalLubricantes1v = roundMoney(totalLubricantes1v);
  totalMo1v = roundMoney(totalMo1v);
  const totalPreventivo1v = roundMoney(totalRepuestos1v + totalLubricantes1v + totalMo1v);

  return {
    error: false,
    modo: 'rango',
    tablaDetalle,
    totalRepuestos1v,
    totalLubricantes1v,
    totalMo1v,
    totalPreventivo1v,
    nVehiculos,
    totalRepuestos: roundMoney(totalRepuestos1v * nVehiculos),
    totalLubricantes: roundMoney(totalLubricantes1v * nVehiculos),
    totalMo: roundMoney(totalMo1v * nVehiculos),
    totalPreventivo: roundMoney(totalPreventivo1v * nVehiculos)
  };
}

/**
 * Saves a generated proforma into the database for auditing & persistence.
 */
export async function guardarProforma(proformaData) {
  try {
    const { data, error } = await supabase
      .from('proformas_generadas')
      .upsert({
        numero_proforma: proformaData.numeroProforma,
        razon_social: proformaData.razonSocial,
        ruc: proformaData.ruc,
        direccion: proformaData.direccion,
        contacto: proformaData.contacto,
        telefono: proformaData.telefono,
        correo: proformaData.correo,
        objeto_contrato: proformaData.objetoContrato,
        modelo_vehiculo: proformaData.modeloVehiculo,
        cantidad_vehiculos: parseInt(proformaData.cantidadVehiculos, 10) || 1,
        km_desde: parseInt(proformaData.kmDesde, 10) || 0,
        km_hasta: parseInt(proformaData.kmHasta, 10) || 0,
        incluye_correctivos: !!proformaData.incluirCorrectivos,
        total_repuestos: proformaData.totalRepuestos,
        total_lubricantes: proformaData.totalLubricantes,
        total_mano_obra: proformaData.totalManoObra,
        total_preventivo: proformaData.totalPreventivo,
        total_correctivo: proformaData.totalCorrectivo,
        gran_total: proformaData.granTotal,
        placas: proformaData.placas || [],
        detalle_km: proformaData.tablaDetalle || []
      }, { onConflict: 'numero_proforma' })
      .select();

    if (error) throw error;
    return { data, error: null };
  } catch (err) {
    console.error("Error saving proforma to Supabase:", err);
    return { data: null, error: err.message };
  }
}

/**
 * Fetches recent proformas generated.
 */
export async function getHistorialProformas() {
  try {
    const { data, error } = await supabase
      .from('proformas_generadas')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(20);

    if (error) throw error;
    return { data: data || [], error: null };
  } catch (err) {
    console.error("Error fetching historial proformas:", err);
    return { data: [], error: err.message };
  }
}
