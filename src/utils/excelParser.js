import * as XLSX from 'xlsx';
import JSZip from 'jszip';

/**
 * Searches for a value in a column next to a label match or list of label aliases.
 * Resilient to different column layouts and capitalization.
 */
export function buscarValor(grid, etiquetas, cols = [2, 3, 4, 5, 6, 1, 0]) {
  if (!grid || !Array.isArray(grid)) return null;
  const etiqList = Array.isArray(etiquetas) ? etiquetas : [etiquetas];
  
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    
    // Check all columns in this row to find the label
    for (let c = 0; c < Math.min(row.length, 6); c++) {
      const cellText = row[c]?.toString().trim().toLowerCase() || "";
      if (!cellText) continue;
      
      for (const etiq of etiqList) {
        const etiqNorm = etiq.toLowerCase().trim();
        if (cellText === etiqNorm || cellText.startsWith(etiqNorm) || cellText.includes(etiqNorm)) {
          // Found the label cell! First look in adjacent columns to the right
          for (let valCol = c + 1; valCol < row.length; valCol++) {
            const v = row[valCol]?.toString().trim();
            if (v && !["nan", "none", "", "null", "undefined", "-", ":"].includes(v.toLowerCase())) {
              return v;
            }
          }
          // Fallback to check default columns
          for (const valCol of cols) {
            if (valCol !== c && valCol < row.length) {
              const v = row[valCol]?.toString().trim();
              if (v && !["nan", "none", "", "null", "undefined", "-", ":"].includes(v.toLowerCase())) {
                return v;
              }
            }
          }
        }
      }
    }
  }
  return null;
}

/**
 * Parses the KM Range: "Desde:" and "Hasta:" from any cell position.
 */
export function parseRangoKM(grid) {
  let kmDesde = null;
  let kmHasta = null;

  if (!grid || !Array.isArray(grid)) return { kmDesde: 5000, kmHasta: 120000 };

  // First pass: look for "Desde" and "Hasta" in any cell
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < Math.min(row.length, 6); c++) {
      const cellText = row[c]?.toString().trim().toLowerCase() || "";
      
      // Look for Desde
      if (cellText === "desde:" || cellText === "desde" || cellText.startsWith("desde")) {
        for (let vc = c + 1; vc < Math.min(row.length, c + 5); vc++) {
          const rawVal = row[vc];
          if (rawVal !== undefined && rawVal !== null && rawVal !== '') {
            const cleanVal = rawVal.toString().replace(/km/gi, '').replace(/\./g, '').replace(/,/g, '').trim();
            const parsed = parseInt(cleanVal, 10);
            if (!isNaN(parsed) && parsed >= 0) {
              kmDesde = parsed;
              break;
            }
          }
        }
      }

      // Look for Hasta
      if (cellText === "hasta:" || cellText === "hasta" || cellText.startsWith("hasta")) {
        for (let vc = c + 1; vc < Math.min(row.length, c + 5); vc++) {
          const rawVal = row[vc];
          if (rawVal !== undefined && rawVal !== null && rawVal !== '') {
            const cleanVal = rawVal.toString().replace(/km/gi, '').replace(/\./g, '').replace(/,/g, '').trim();
            const parsed = parseInt(cleanVal, 10);
            if (!isNaN(parsed) && parsed > 0) {
              kmHasta = parsed;
              break;
            }
          }
        }
      }
    }
  }

  // Fallback pass: look for "incluir en la cotización" or "mantenimientos"
  if (!kmDesde || !kmHasta) {
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      const rowStr = row.map(v => (v || '').toString().toLowerCase()).join(' ');
      if (rowStr.includes("incluir en la cotización") || rowStr.includes("mantenimientos que se deben") || rowStr.includes("mantenimientos")) {
        for (let vc = 0; vc < row.length; vc++) {
          const rawVal = row[vc];
          if (rawVal !== undefined && rawVal !== null && rawVal !== '') {
            const cleanVal = rawVal.toString().replace(/km/gi, '').replace(/\./g, '').replace(/,/g, '').trim();
            const parsed = parseInt(cleanVal, 10);
            if (!isNaN(parsed) && parsed >= 1000 && !kmDesde) {
              kmDesde = parsed;
              break;
            }
          }
        }

        // Search for "Hasta" in the next 5 rows
        for (let j = r + 1; j < Math.min(r + 6, grid.length); j++) {
          const r2 = grid[j] || [];
          for (let c = 0; c < r2.length; c++) {
            const ct = (r2[c] || '').toString().toLowerCase();
            if (ct.includes("hasta")) {
              for (let vc = c + 1; vc < r2.length; vc++) {
                const rawVal2 = r2[vc];
                if (rawVal2 !== undefined && rawVal2 !== null && rawVal2 !== '') {
                  const cleanVal2 = rawVal2.toString().replace(/km/gi, '').replace(/\./g, '').replace(/,/g, '').trim();
                  const parsed2 = parseInt(cleanVal2, 10);
                  if (!isNaN(parsed2) && parsed2 > 0) {
                    kmHasta = parsed2;
                    break;
                  }
                }
              }
            }
          }
          if (kmHasta) break;
        }
        break;
      }
    }
  }

  return { kmDesde: kmDesde || 5000, kmHasta: kmHasta || 120000 };
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
  if (!grid || !Array.isArray(grid)) return placas;
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < Math.min(row.length, 6); c++) {
      const cellText = row[c]?.toString().trim().toUpperCase() || "";
      if (cellText === "PLACAS" || cellText === "PLACA" || cellText.startsWith("PLACA")) {
        for (let vc = c + 1; vc < row.length; vc++) {
          if (row[vc] !== undefined && row[vc] !== null) {
            const v = row[vc].toString().trim();
            if (v && !["nan", "placas", "placa", "none", "", "-"].includes(v.toLowerCase())) {
              const splitPlacas = v.split(/[,;/]+/).map(p => p.trim().toUpperCase()).filter(Boolean);
              placas.push(...splitPlacas);
            }
          }
        }
      }
    }
  }
  return Array.from(new Set(placas));
}

/**
 * Parses individual vehicle maintenances from Observation or free-form text.
 * Handles patterns like:
 * "HEI-1859= MANTTO DE 35.000 KM"
 * "HEI-1845= MANTTO DE 30.000 KM"
 * "HEI-1862= MANTTO DE 35.000 KM"
 * "ABC-1234: 50.000 KM"
 * "10000 km, 20000 km"
 */
export function parseMantenimientosObservacion(obsText) {
  if (!obsText || typeof obsText !== 'string') return [];
  const lines = obsText.split(/[\r\n]+/);
  const items = [];

  for (const line of lines) {
    const lineClean = line.trim();
    if (!lineClean) continue;

    // 1. Extract plate if present (e.g. HEI-1859, PBA-1234, etc.)
    const plateMatch = /([A-Za-z]{3}-?\d{3,4})/i.exec(lineClean);
    const plate = plateMatch ? plateMatch[1].toUpperCase() : null;

    // 2. Strip plate from text to avoid numeric plate digits being mistaken for KM
    let textWithoutPlate = lineClean;
    if (plate) {
      textWithoutPlate = textWithoutPlate.replace(new RegExp(plate, 'i'), '');
    }

    // 3. Match KM
    let kmVal = null;

    // Pattern A: number followed by KM (e.g. 35.000 KM, 35000km)
    const kmPatternA = /(\d+(?:[\.,]\d+)*)\s*KM/i.exec(textWithoutPlate);
    if (kmPatternA) {
      const raw = kmPatternA[1].replace(/\./g, '').replace(/,/g, '');
      const num = parseInt(raw, 10);
      if (!isNaN(num) && num >= 5000 && num <= 300000) {
        kmVal = num;
      }
    }

    // Pattern B: number after MANTTO, MANTENIMIENTO, =, or :
    if (!kmVal) {
      const kmPatternB = /(?:MANTTO|MANTTO\.|MANTENIMIENTO|=|:)\s*(?:DE\s*)?(\d+(?:[\.,]\d+)*)/i.exec(textWithoutPlate);
      if (kmPatternB) {
        const raw = kmPatternB[1].replace(/\./g, '').replace(/,/g, '');
        const num = parseInt(raw, 10);
        if (!isNaN(num) && num >= 5000 && num <= 300000) {
          kmVal = num;
        }
      }
    }

    // Pattern C: any 4 to 6 digit number >= 5000 and <= 300000
    if (!kmVal) {
      const allNumbers = textWithoutPlate.match(/\b\d[\d\.,]*\b/g);
      if (allNumbers) {
        for (const n of allNumbers) {
          const raw = n.replace(/\./g, '').replace(/,/g, '');
          const num = parseInt(raw, 10);
          if (!isNaN(num) && num >= 5000 && num <= 300000) {
            kmVal = num;
            break;
          }
        }
      }
    }

    if (kmVal) {
      items.push({
        id: `item-${items.length + 1}`,
        placa: plate || `Vehículo ${items.length + 1}`,
        km: kmVal,
        raw: lineClean
      });
    }
  }

  return items;
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
