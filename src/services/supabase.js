import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || "https://mblsjoreokpsjtdzjrae.supabase.co";
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || "";

export const supabase = createClient(supabaseUrl, supabaseAnonKey || "dummy-anon-key-for-build");

/**
 * Fetches all vehicle models in the database.
 */
export async function getModelosVehiculo() {
  try {
    const { data, error } = await supabase
      .from('modelos_vehiculo')
      .select('*, planes_mantenimiento(*)')
      .order('marca', { ascending: true })
      .order('modelo', { ascending: true });
      
    if (error) throw error;
    return { data, error: null };
  } catch (err) {
    console.error("Error fetching modelos_vehiculo:", err);
    return { data: [], error: err.message };
  }
}

/**
 * Searches for the best matching plan in Supabase given a free-text model string
 * (e.g. from the client's Excel quotation: "WINGLE 7 4X4", "POER DIESEL", "TANK 300", etc.)
 */
export async function buscarPlanPorModelo(modeloTexto = "") {
  try {
    const normalized = (modeloTexto || "").toUpperCase().trim();
    const { data: modelos, error } = await supabase
      .from('modelos_vehiculo')
      .select('*, planes_mantenimiento(*)');

    if (error) throw error;
    if (!modelos || modelos.length === 0) return { plan: null, modelo: null };

    // Ranking algorithm for matching
    let bestMatch = null;
    let highestScore = -1;

    for (const m of modelos) {
      let score = 0;
      const mName = m.nombre_completo.toUpperCase();
      const mModelo = m.modelo.toUpperCase();
      const mTraccion = (m.traccion || '').toUpperCase();
      const mMotor = (m.motor || '').toUpperCase();

      // Check for vehicle model keywords
      if (normalized.includes(mModelo)) score += 10;
      if (mModelo.includes("WINGLE 7") && (normalized.includes("W7") || normalized.includes("WINGLE 7") || normalized.includes("WINGLE7"))) score += 12;
      if (mModelo.includes("POER") && normalized.includes("POER")) score += 12;
      if (mModelo.includes("TANK 300") && (normalized.includes("TANK 300") || normalized.includes("TANK300"))) score += 15;
      if (mModelo.includes("TANK 500") && (normalized.includes("TANK 500") || normalized.includes("TANK500"))) score += 15;
      if (mModelo.includes("KYC") && (normalized.includes("KYC") || normalized.includes("F3"))) score += 12;
      if (mModelo.includes("WINGLE 2.8") && (normalized.includes("2.8") || normalized.includes("2,8"))) score += 15;

      // Check traction
      if (normalized.includes("4X4") && (mTraccion.includes("4X4") || mName.includes("4X4"))) score += 5;
      if (normalized.includes("4X2") && (mTraccion.includes("4X2") || mName.includes("4X2"))) score += 5;

      // Check motor
      if (normalized.includes("DIESEL") && (mMotor.includes("DIESEL") || mName.includes("DIESEL"))) score += 3;
      if (normalized.includes("GASOLINA") && (mMotor.includes("GASOLINA") || mName.includes("GASOLINA"))) score += 3;
      if (normalized.includes("POLIC") && mName.includes("POLIC")) score += 8;

      if (score > highestScore && score > 0) {
        highestScore = score;
        bestMatch = m;
      }
    }

    if (bestMatch && bestMatch.planes_mantenimiento && bestMatch.planes_mantenimiento.length > 0) {
      return { plan: bestMatch.planes_mantenimiento[0], modelo: bestMatch };
    }

    return { plan: null, modelo: null };
  } catch (err) {
    console.error("Error matching plan:", err);
    return { plan: null, modelo: null, error: err.message };
  }
}

/**
 * Fetches the maintenance costs for a given plan and KM range.
 */
export async function getCostosPorKm(planId, kmDesde, kmHasta, cantidadVehiculos = 1) {
  try {
    const kDesde = parseInt(kmDesde, 10) || 0;
    const kHasta = parseInt(kmHasta, 10) || 250000;
    const nVehiculos = parseInt(cantidadVehiculos, 10) || 1;

    const { data, error } = await supabase
      .from('plan_costos_km')
      .select('*')
      .eq('plan_id', planId)
      .gte('km', kDesde)
      .lte('km', kHasta)
      .order('km', { ascending: true });

    if (error) throw error;

    if (!data || data.length === 0) {
      const { data: allKms } = await supabase
        .from('plan_costos_km')
        .select('km')
        .eq('plan_id', planId)
        .order('km', { ascending: true });

      const kmsDisp = allKms ? allKms.map(x => x.km) : [];
      return {
        error: true,
        kmsDisponibles: kmsDisp,
        message: `No hay mantenimientos en el rango [${kDesde} - ${kHasta} KM]. Kilómetros disponibles: ${kmsDisp.join(', ')}`
      };
    }

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

      return {
        km: row.km,
        repuestos: rep,
        lubricantes: lub,
        mano_obra: mo,
        total: tot
      };
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
  } catch (err) {
    console.error("Error fetching costos por km:", err);
    return { error: true, message: err.message };
  }
}

/**
 * Fetches corrective items catalog for reference.
 */
export async function getCatalogoCorrectivo(modeloReferencia = "") {
  try {
    let query = supabase.from('catalogo_correctivo').select('*');
    if (modeloReferencia) {
      query = query.ilike('modelo_referencia', `%${modeloReferencia}%`);
    }
    const { data, error } = await query.order('item_numero', { ascending: true });
    if (error) throw error;
    return { data: data || [], error: null };
  } catch (err) {
    console.error("Error fetching catalogo_correctivo:", err);
    return { data: [], error: err.message };
  }
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
