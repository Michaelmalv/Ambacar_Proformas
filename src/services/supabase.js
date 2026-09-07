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
export async function buscarPlanPorModelo(modeloTexto = "") {
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
    if (normalized === mName || normalized === mModelo) score += 100;
    if (mName.includes(normalized) && normalized.length > 3) score += 40;
    if (normalized.includes(mName)) score += 40;

    // Model families
    if (mModelo.includes("WINGLE 7") && (normalized.includes("W7") || normalized.includes("WINGLE 7") || normalized.includes("WINGLE7") || normalized.includes("WINGLE"))) score += 30;
    if (mModelo.includes("POER") && normalized.includes("POER")) score += 30;
    if (mModelo.includes("TANK 300") && (normalized.includes("TANK 300") || normalized.includes("TANK300") || (normalized.includes("TANK") && normalized.includes("300")))) score += 40;
    if (mModelo.includes("TANK 500") && (normalized.includes("TANK 500") || normalized.includes("TANK500") || (normalized.includes("TANK") && normalized.includes("500")))) score += 40;
    if (mModelo.includes("KYC") && (normalized.includes("KYC") || normalized.includes("F3"))) score += 30;
    if (mModelo.includes("WINGLE 2.8") && (normalized.includes("2.8") || normalized.includes("2,8"))) score += 40;

    // Traction matching (4x4 vs 4x2)
    if ((normalized.includes("4X4") || normalized.includes("4 X 4") || normalized.includes("4WD") || normalized.includes("DOBLE TRACCION")) && (mTraccion.includes("4X4") || mName.includes("4X4"))) score += 20;
    if ((normalized.includes("4X2") || normalized.includes("4 X 2") || normalized.includes("2WD") || normalized.includes("SIMPLE")) && (mTraccion.includes("4X2") || mName.includes("4X2"))) score += 20;

    // Fuel & engine matching
    if (normalized.includes("DIESEL") && (mMotor.includes("DIESEL") || mName.includes("DIESEL"))) score += 10;
    if (normalized.includes("GASOLINA") && (mMotor.includes("GASOLINA") || mName.includes("GASOLINA"))) score += 10;
    if ((normalized.includes("POLIC") || normalized.includes("PATRULL")) && mName.includes("POLIC")) score += 25;

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

/**
 * Calculates maintenance costs for a given plan and KM range.
 */
export async function getCostosPorKm(planId, kmDesde, kmHasta, cantidadVehiculos = 1) {
  const kDesde = parseInt(kmDesde, 10) || 0;
  const kHasta = parseInt(kmHasta, 10) || 250000;
  const nVehiculos = parseInt(cantidadVehiculos, 10) || 1;

  // 1. Check if Supabase has data
  try {
    const { data, error } = await supabase
      .from('plan_costos_km')
      .select('*')
      .eq('plan_id', planId)
      .gte('km', kDesde)
      .lte('km', kHasta)
      .order('km', { ascending: true });

    if (!error && data && data.length > 0) {
      let totalRepuestos1v = 0;
      let totalLubricantes1v = 0;
      let totalMo1v = 0;

      const tablaDetalle = data.map(row => {
        const rep = parseFloat(row.total_repuestos) || 0;
        const lub = parseFloat(row.total_lubricantes) || 0;
        const mo = parseFloat(row.total_mano_obra) || 0;
        const tot = parseFloat(row.total_km) || (rep + lub + mo);

        totalRepuestos1v += rep;
        totalLubricantes1v += lub;
        totalMo1v += mo;

        return { km: row.km, repuestos: rep, lubricantes: lub, mano_obra: mo, total: tot };
      });

      const totalPreventivo1v = totalRepuestos1v + totalLubricantes1v + totalMo1v;

      return {
        error: false,
        tablaDetalle,
        totalRepuestos1v,
        totalLubricantes1v,
        totalMo1v,
        totalPreventivo1v,
        nVehiculos,
        totalRepuestos: totalRepuestos1v * nVehiculos,
        totalLubricantes: totalLubricantes1v * nVehiculos,
        totalMo: totalMo1v * nVehiculos,
        totalPreventivo: totalPreventivo1v * nVehiculos
      };
    }
  } catch (err) {
    console.warn("Supabase query failed, falling back to local dataset:", err);
  }

  // 2. Use Bundled Local Dataset from 5 Excels
  const localPlan = PLANES_MANTENIMIENTO_DATA.find(p => p.id === planId || p.codigo_plan === planId);
  if (!localPlan) {
    return {
      error: true,
      message: `No se encontró el plan de mantenimiento seleccionado.`
    };
  }

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

  const totalPreventivo1v = totalRepuestos1v + totalLubricantes1v + totalMo1v;

  return {
    error: false,
    tablaDetalle,
    totalRepuestos1v,
    totalLubricantes1v,
    totalMo1v,
    totalPreventivo1v,
    nVehiculos,
    totalRepuestos: totalRepuestos1v * nVehiculos,
    totalLubricantes: totalLubricantes1v * nVehiculos,
    totalMo: totalMo1v * nVehiculos,
    totalPreventivo: totalPreventivo1v * nVehiculos
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
