import * as XLSX from 'xlsx';
import JSZip from 'jszip';

/**
 * Searches for a value in a column next to a label match.
 * Matches python's `buscar_valor(df, etiqueta, cols=[3,4,5,2])`
 */
export function buscarValor(grid, etiqueta, cols = [3, 4, 5, 2]) {
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    const c1 = row[1]?.toString().trim() || "";
    if (c1.toLowerCase().includes(etiqueta.toLowerCase())) {
      for (const c of cols) {
        if (c < row.length) {
          const v = row[c]?.toString().trim();
          if (v && !["nan", "none", ""].includes(v.toLowerCase())) {
            return v;
          }
        }
      }
    }
  }
  return null;
}

/**
 * Parses the KM Range: "Desde:" and "Hasta:"
 */
export function parseRangoKM(grid) {
  let kmDesde = null;
  let kmHasta = null;

  // First pass: look for "Desde:" and "Hasta:" cells
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    for (const cc of [1, 3]) {
      if (cc >= row.length) continue;
      const celda = row[cc]?.toString().trim() || "";
      if (celda === "Desde:") {
        for (const vc of [cc + 1, 4, 5]) {
          if (vc < row.length && row[vc] !== undefined && row[vc] !== null) {
            const cleanVal = row[vc].toString().replace(/,/g, '');
            const parsed = parseInt(parseFloat(cleanVal), 10);
            if (!isNaN(parsed)) {
              kmDesde = parsed;
              break;
            }
          }
        }
      }
      if (celda === "Hasta:") {
        for (const vc of [cc + 1, 4, 5]) {
          if (vc < row.length && row[vc] !== undefined && row[vc] !== null) {
            const cleanVal = row[vc].toString().replace(/,/g, '');
            const parsed = parseInt(parseFloat(cleanVal), 10);
            if (!isNaN(parsed)) {
              kmHasta = parsed;
              break;
            }
          }
        }
      }
    }
  }

  // Fallback pass: look for "incluir en la cotización"
  if (!kmDesde || !kmHasta) {
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      const c1 = row[1]?.toString().trim() || "";
      if (c1.toLowerCase().includes("incluir en la cotización") || c1.toLowerCase().includes("mantenimientos que se deben")) {
        for (const vc of [3, 4, 5]) {
          if (vc < row.length && row[vc] !== undefined && row[vc] !== null) {
            const cleanVal = row[vc].toString().replace(/,/g, '');
            const parsed = parseInt(parseFloat(cleanVal), 10);
            if (!isNaN(parsed) && parsed > 1000 && !kmDesde) {
              kmDesde = parsed;
              break;
            }
          }
        }

        // Search for "Hasta" in the next 5 rows
        for (let j = r + 1; j < Math.min(r + 5, grid.length); j++) {
          const r2 = grid[j] || [];
          const c3 = r2[3]?.toString().trim() || "";
          if (c3.includes("Hasta") && r2[4] !== undefined && r2[4] !== null) {
            const cleanVal = r2[4].toString().replace(/,/g, '');
            const parsed = parseInt(parseFloat(cleanVal), 10);
            if (!isNaN(parsed)) {
              kmHasta = parsed;
              break;
            }
          }
        }
        break;
      }
    }
  }

  return { kmDesde, kmHasta };
}

/**
 * Detects correctivo checkboxes by reading drawings inside the ZIP.
 * Matches python's `detectar_correctivos_excel(xlsx_bytes)`
 */
export async function detectarCorrectivosExcel(fileBuffer, grid) {
  try {
    const zip = await JSZip.loadAsync(fileBuffer);
    
    // Check if VML and drawing files exist
    const vmlPath = 'xl/drawings/vmlDrawing1.vml';
    const drawingPath = 'xl/drawings/drawing1.xml';
    
    if (!zip.files[vmlPath] || !zip.files[drawingPath]) {
      return null;
    }
    
    const vml = await zip.files[vmlPath].async('text');
    const drawing = await zip.files[drawingPath].async('text');
    
    // Find twoCellAnchors in drawing1.xml
    const anchors = [];
    const anchorRegex = /<xdr:twoCellAnchor[^>]*>([\s\S]*?)<\/xdr:twoCellAnchor>/g;
    let match;
    while ((match = anchorRegex.exec(drawing)) !== null) {
      anchors.push(match[1]);
    }
    
    const anchorRows = [];
    for (const a of anchors) {
      const rowMatch = /<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(a);
      const colMatch = /<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/.exec(a);
      if (rowMatch && colMatch) {
        anchorRows.push({
          row: parseInt(rowMatch[1], 10) + 1,
          col: parseInt(colMatch[1], 10) + 1
        });
      }
    }
    
    // Find shapes in vmlDrawing1.vml
    const shapesRaw = [];
    const shapeRegex = /<v:shape id="(?!_x0000_t)[^"]*"[\s\S]*?<\/v:shape>/g;
    let sm;
    while ((sm = shapeRegex.exec(vml)) !== null) {
      shapesRaw.push(sm[0]);
    }
    
    const shapes = shapesRaw.map(s => {
      const mtMatch = /margin-top:\s*([.\d]+)pt/.exec(s);
      const marginTop = mtMatch ? parseFloat(mtMatch[1]) : 0;
      const checked = s.includes('<x:Checked>');
      return { marginTop, checked };
    }).sort((a, b) => a.marginTop - b.marginTop);
    
    // Map shapes to anchors and search for the correctivo row
    for (let idx = 0; idx < Math.min(shapes.length, anchorRows.length); idx++) {
      const { row } = anchorRows[idx];
      const b = grid[row - 1]?.[1]?.toString().trim() || "";
      if (b.toLowerCase().includes("correctivo") && b.toLowerCase().includes("incluir")) {
        const siChecked = shapes[idx].checked;
        const noChecked = shapes[idx + 1] ? shapes[idx + 1].checked : false;
        return { siChecked, noChecked };
      }
    }
  } catch (err) {
    console.error("VML checkbox parsing error:", err);
  }
  return null;
}

/**
 * Parses plate numbers from the spreadsheet.
 */
export function parsePlacas(grid) {
  const placas = [];
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    const c1 = row[1]?.toString().trim() || "";
    if (c1.toUpperCase() === "PLACAS") {
      for (const vc of [3, 4, 5]) {
        if (vc < row.length && row[vc] !== undefined && row[vc] !== null) {
          const v = row[vc].toString().trim();
          if (v && !["nan", "placas", ""].includes(v.toLowerCase())) {
            placas.push(v);
            break;
          }
        }
      }
    }
  }
  return placas;
}

/**
 * Helper to convert text (like "10.000 KM") into a numeric integer.
 */
export function kmAInt(s) {
  if (s === undefined || s === null) return null;
  const str = s.toString().toUpperCase().replace(/KM/g, "").replace(/\./g, "").replace(/,/g, "").trim();
  const parsed = parseInt(str, 10);
  return isNaN(parsed) ? null : parsed;
}

/**
 * Calculates the totals and list details of maintenance based on grid data and KM range.
 */
export function calcularPlanMantenimiento(planGrid, kmDesde, kmHasta, cantidadVehiculos) {
  let filaEnc = null;
  for (let r = 0; r < planGrid.length; r++) {
    const c0 = planGrid[r]?.[0]?.toString().trim() || "";
    if (c0.toUpperCase().includes("GWM") && c0.toUpperCase().includes("DIESEL")) {
      filaEnc = r;
      break;
    }
  }
  
  if (filaEnc === null) {
    filaEnc = 1;
  }
  
  // Find boundaries of this block
  let finBloque = planGrid.length;
  for (let r = filaEnc + 5; r < planGrid.length; r++) {
    const c0 = planGrid[r]?.[0]?.toString().trim() || "";
    if (c0.toUpperCase().includes("GWM") && c0.toUpperCase().includes("DIESEL")) {
      finBloque = r - 1;
      break;
    }
  }
  
  const filaKmRow = planGrid[filaEnc] || [];
  const colsRango = [];
  for (let ci = 0; ci < filaKmRow.length; ci++) {
    const v = filaKmRow[ci];
    const km = kmAInt(v);
    if (km && km >= kmDesde && km <= kmHasta) {
      colsRango.push({ colIndex: ci, km });
    }
  }
  
  if (colsRango.length === 0) {
    const kmsDisponibles = [];
    for (let ci = 0; ci < filaKmRow.length; ci++) {
      const km = kmAInt(filaKmRow[ci]);
      if (km) kmsDisponibles.push(km);
    }
    return {
      error: true,
      kmsDisponibles: Array.from(new Set(kmsDisponibles)).sort((a, b) => a - b),
    };
  }
  
  // Find rows of totals within the block
  let filaRep = null;
  let filaLub = null;
  let filaMo = null;
  let filaTot = null;
  for (let r = filaEnc; r < finBloque; r++) {
    const etiq = planGrid[r]?.[0]?.toString().trim().toUpperCase() || "";
    if (etiq === "TOTAL REPUESTOS" && filaRep === null) filaRep = r;
    if (etiq === "TOTAL LUBRICANTES" && filaLub === null) filaLub = r;
    if (etiq === "TOTAL M/O" && filaMo === null) filaMo = r;
    if (etiq === "TOTAL MANTENIMIENTO X KM" && filaTot === null) filaTot = r;
  }
  
  const gv = (r, c) => {
    if (r === null || r === undefined) return 0.0;
    const val = planGrid[r]?.[c];
    if (val === undefined || val === null || val === '') return 0.0;
    const num = parseFloat(val.toString().replace(/,/g, ''));
    return isNaN(num) ? 0.0 : num;
  };
  
  let totalRepuestos1v = 0;
  let totalLubricantes1v = 0;
  let totalMo1v = 0;
  
  colsRango.forEach(({ colIndex }) => {
    totalRepuestos1v += gv(filaRep, colIndex);
    totalLubricantes1v += gv(filaLub, colIndex);
    totalMo1v += gv(filaMo, colIndex);
  });
  
  const totalPreventivo1v = totalRepuestos1v + totalLubricantes1v + totalMo1v;
  
  // Details
  const tablaDetalle = colsRango.map(({ colIndex, km }) => {
    const rep = gv(filaRep, colIndex);
    const lub = gv(filaLub, colIndex);
    const mo = gv(filaMo, colIndex);
    const tot = filaTot !== null ? gv(filaTot, colIndex) : (rep + lub + mo);
    return {
      km: km,
      repuestos: rep,
      lubricantes: lub,
      mano_obra: mo,
      total: tot
    };
  });
  
  const nVehiculos = parseInt(cantidadVehiculos, 10) || 1;
  const totalRepuestos = totalRepuestos1v * nVehiculos;
  const totalLubricantes = totalLubricantes1v * nVehiculos;
  const totalMo = totalMo1v * nVehiculos;
  const totalPreventivo = totalPreventivo1v * nVehiculos;
  
  return {
    error: false,
    colsRango,
    totalRepuestos1v,
    totalLubricantes1v,
    totalMo1v,
    totalPreventivo1v,
    tablaDetalle,
    nVehiculos,
    totalRepuestos,
    totalLubricantes,
    totalMo,
    totalPreventivo,
  };
}
