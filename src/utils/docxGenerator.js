import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  ImageRun, 
  Header, 
  Footer,
  AlignmentType, 
  WidthType, 
  VerticalAlign,
  BorderStyle
} from 'docx';

// Convert base64 data URL to Uint8Array for docx ImageRun
function base64ToUint8Array(base64String) {
  const cleanBase64 = base64String.replace(/^data:image\/\w+;base64,/, "");
  const binaryString = window.atob(cleanBase64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

// Spacing helper: Converts pt to dxa (1 pt = 20 dxa)
const pt = (v) => v * 20;

// Width helper: Converts cm to dxa (1 cm = 567 dxa)
const cm = (v) => Math.round(v * 567);

export async function generarProformaDocx(data) {
  const {
    razonSocial = "",
    ruc = "",
    direccion = "",
    contacto = "",
    telefono = "",
    correo = "",
    modeloVehiculo = "",
    cantidadVehiculos = 1,
    plazoEjecucion = "365",
    vigenciaOferta = "90 días",
    observacion = "",
    objetoContrato = "Plan de mantenimiento preventivo y correctivo de vehículos",
    placas = [],
    incluirCorrectivos = false,
    hojaElegida = "",
    
    // Totals
    totalRepuestos = 0,
    totalLubricantes = 0,
    totalManoObra = 0,
    totalPreventivo = 0,
    totalCorrectivo = 0,
    granTotal = 0,
    
    // Images (base64 data URLs)
    headerImage = null,
    footerImage = null
  } = data;

  const today = new Date();
  const year = today.getFullYear();
  const monthNames = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
  ];
  const fechaStr = `${today.getDate()} de ${monthNames[today.getMonth()]} de ${year}`;
  
  const mmStr = String(today.getMonth() + 1).padStart(2, '0');
  const ddStr = String(today.getDate()).padStart(2, '0');
  const numProforma = `AMB-PV-${year}-AUTO-${mmStr}${ddStr}`;

  // Paragraph helper
  const p = (text, options = {}) => {
    const runs = [];
    if (text) {
      runs.push(new TextRun({
        text,
        bold: !!options.bold,
        italic: !!options.italic,
        size: options.size || 18, // default 9pt (size 18)
        color: options.color || undefined,
        font: "Arial"
      }));
    }
    
    if (options.runs) {
      options.runs.forEach(r => {
        runs.push(new TextRun({
          text: r.text || "",
          bold: !!r.bold,
          italic: !!r.italic,
          size: r.size || options.size || 18,
          color: r.color || undefined,
          font: "Arial"
        }));
      });
    }

    return new Paragraph({
      alignment: options.align || AlignmentType.JUSTIFY,
      spacing: {
        before: pt(options.before || 0),
        after: pt(options.after || 4),
      },
      indent: options.leftIndent ? { left: cm(options.leftIndent) } : undefined,
      border: options.border || undefined,
      children: runs
    });
  };

  // Heading helper
  const heading = (text, spaceBefore = 6, spaceAfter = 3) => {
    return p(text, {
      bold: true,
      size: 20, // 10pt
      color: "CC0000",
      align: AlignmentType.LEFT,
      before: spaceBefore,
      after: spaceAfter
    });
  };

  // Table Cell helper
  const cell = (text, options = {}) => {
    let cellChildren = [];
    if (options.children) {
      cellChildren = options.children;
    } else {
      const runs = [];
      if (text) {
        runs.push(new TextRun({
          text,
          bold: !!options.bold,
          italic: !!options.italic,
          size: options.size || 18, // 9pt
          color: options.color || undefined,
          font: "Arial"
        }));
      }
      if (options.runs) {
        options.runs.forEach(r => {
          runs.push(new TextRun({
            text: r.text || "",
            bold: !!r.bold,
            italic: !!r.italic,
            size: r.size || options.size || 18,
            color: r.color || undefined,
            font: "Arial"
          }));
        });
      }
      cellChildren.push(new Paragraph({
        alignment: options.align || AlignmentType.LEFT,
        spacing: { before: pt(2), after: pt(2) },
        children: runs
      }));
    }

    return new TableCell({
      children: cellChildren,
      shading: options.bg ? { fill: options.bg } : undefined,
      width: options.widthCm ? { size: cm(options.widthCm), type: WidthType.DXA } : undefined,
      columnSpan: options.columnSpan || undefined,
      rowSpan: options.rowSpan || undefined,
      verticalAlign: VerticalAlign.CENTER
    });
  };

  // Header image setup
  const headerParagraphs = [];
  if (headerImage) {
    try {
      headerParagraphs.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 0 },
        children: [
          new ImageRun({
            data: base64ToUint8Array(headerImage),
            transformation: {
              width: 642, // ~17cm in pixels at 96dpi
              height: 70,  // aspect ratio adjusted
            }
          })
        ]
      }));
    } catch (e) {
      console.error("Error drawing header image in docx:", e);
      headerParagraphs.push(new Paragraph("AMBACAR CIA. LTDA."));
    }
  } else {
    headerParagraphs.push(new Paragraph("AMBACAR CIA. LTDA."));
  }

  // Footer image setup
  const footerParagraphs = [];
  if (footerImage) {
    try {
      footerParagraphs.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0 },
        children: [
          new ImageRun({
            data: base64ToUint8Array(footerImage),
            transformation: {
              width: 642,
              height: 50,
            }
          })
        ]
      }));
    } catch (e) {
      console.error("Error drawing footer image in docx:", e);
      footerParagraphs.push(new Paragraph("AMBACAR CIA. LTDA. - Contratación Pública"));
    }
  } else {
    footerParagraphs.push(new Paragraph("AMBACAR CIA. LTDA. - Contratación Pública"));
  }

  // Formatting currency helper
  const fmt = (val) => {
    return new Intl.NumberFormat('en-US', {
      style: 'decimal',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(val);
  };

  // --- BUILD THE DOCUMENT CONTENT ---
  const docChildren = [
    // Fechas and Nro Proforma
    p(`Quito, ${fechaStr}`, { align: AlignmentType.RIGHT, after: 0, size: 18 }),
    p(`Proforma Nro. ${numProforma}`, { align: AlignmentType.RIGHT, after: 6, size: 18 }),
    
    // Red horizontal separator
    new Paragraph({
      spacing: { after: pt(8) },
      border: {
        bottom: {
          color: "CC0000",
          space: 2,
          style: BorderStyle.SINGLE,
          size: 12, // ~1.5 pt
        }
      }
    }),

    // Document Title
    p(objetoContrato || "Plan de mantenimiento preventivo y correctivo de vehículos", {
      bold: true,
      size: 26, // 13pt
      color: "CC0000",
      align: AlignmentType.CENTER,
      after: 8
    }),

    // Datos Proveedor
    heading("DATOS DEL PROVEEDOR:"),
    ...[
      ["Razón Social: ", "AMBACAR CIA LTDA."],
      ["RUC: ", "1890010705001"],
      ["Dirección: ", "Av. Indoamérica Km1, Ambato."],
      ["Teléfonos de contacto: ", "0983509888 - 0993758313"],
      ["Persona de Contacto: ", "Geovanna Pilapanta / Christian Salazar / Byron López"],
      ["Correo Electrónico: ", "admincontratacion@ambacar.com / contratacionpublica02@ambacar.com / contratacionpublica04@ambacar.com"]
    ].map(([lbl, val]) => p("", {
      leftIndent: 0.3,
      after: 2,
      runs: [
        { text: "- ", bold: false },
        { text: lbl, bold: true },
        { text: val, bold: false }
      ]
    })),

    new Paragraph({ spacing: { after: pt(4) } }),

    // Datos Cliente Table
    heading("DATOS DEL CLIENTE:"),
    new Table({
      columnWidths: [cm(4), cm(13)],
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" }
      },
      rows: [
        ["Razón Social:", razonSocial],
        ["RUC:", ruc],
        ["Dirección:", direccion],
        ["Teléfono:", telefono],
        ["Correo electrónico:", correo]
      ].map(([lbl, val]) => new TableRow({
        children: [
          cell(lbl, { bold: true, bg: "E0E0E0", widthCm: 4 }),
          cell(val, { widthCm: 13 })
        ]
      }))
    }),

    new Paragraph({ spacing: { after: pt(4) } }),

    // Componentes Ofertados Table
    heading("COMPONENTES OFERTADOS:"),
    new Table({
      columnWidths: [cm(3), cm(14)],
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" }
      },
      rows: [
        new TableRow({
          children: [
            cell("CODIGO CPC", { bold: true, bg: "E0E0E0", align: AlignmentType.CENTER, widthCm: 3 }),
            cell("DESCRIPCIÓN", { bold: true, bg: "E0E0E0", widthCm: 14 })
          ]
        }),
        new TableRow({
          children: [
            cell("87141", { align: AlignmentType.CENTER, widthCm: 3 }),
            cell("SERVICIOS DE MANTENIMIENTO Y REPARACION DE VEHÍCULOS DE MOTOR.", { widthCm: 14 })
          ]
        })
      ]
    }),

    new Paragraph({ spacing: { after: pt(6) } }),

    // Preventivo Legal Texts
    heading("MANTENIMIENTO PREVENTIVO"),
    p("El mencionado mantenimiento conlleva la programación de inspecciones, tanto de funcionamiento como de seguridad, ajustes, reparaciones o cambio de repuestos, análisis, limpieza, cambio de lubricantes, calibración, mano de obra, entre otras, que deben desarrollarse de forma periódica con base en la planificación establecida por el proveedor autorizado AMBACAR CIA. LTDA., conforme se determina a continuación:"),
    p("Considerando la información remitida por el proveedor autorizado para brindar el servicio, este tipo de mantenimiento debe llevarse a cabo considerándose los plazos establecidos para cada vehículo (tiempo, kilometro, recorrido), esto es cada 5.000 Km, sin dejar de mencionar que los mismos se realizan en condiciones normales, conforme el siguiente detalle:"),
    
    ...[
      ["a) ", "El proveedor deberá garantizar el cumplimiento de la garantía técnica en cumplimiento del principio de vigencia tecnológica establecido en la normativa legal vigente, en todos los trabajos desarrollados por el contratista."],
      ["b) ", "El proveedor realizará este mantenimiento en coordinación con el Administrador/a del contrato y tendrá lugar antes de que ocurran las fallas o averías en los vehículos."],
      ["c) ", "El oferente deberá disponer del recurso humano técnico y calificado, maquinaria, equipo y herramientas suficientes y necesarias para la prestación del servicio contratado, en sujeción a estos términos de referencia."],
      ["d) ", "El proveedor designará un funcionario quien conjuntamente con el Administrador/a del contrato coordinará de manera eficiente y eficaz todas las actividades relacionadas con la prestación de servicio, tanto para la solicitud de atención de vehículos, así como la tramitación de las facturas correspondientes."],
      ["e) ", "Los trabajos de mantenimiento preventivo de los vehículos serán efectuados de acuerdo al Plan de Mantenimiento Preventivo y Correctivo otorgado por el proveedor autorizado: AMBACAR Cía. Ltda."]
    ].map(([letra, texto]) => p("", {
      leftIndent: 0.5,
      after: 2,
      runs: [
        { text: letra, bold: true },
        { text: texto }
      ]
    })),

    new Paragraph({ spacing: { after: pt(4) } }),

    // Correctivo Legal Texts
    heading("MANTENIMIENTO CORRECTIVO"),
    ...[
      "El mantenimiento correctivo es aplicado en casos específicos cuando los vehículos sufran desperfectos mecánicos inesperados, debido a imprevistos y/o operación anormal del vehículo por parte del/los conductores y/o usuarios; los cuales no se encuentren cubiertos por la garantía técnica otorgada por el fabricante de las camionetas.",
      "Para la realización de trabajos de mantenimiento correctivos; y establecer el presupuesto de este rubro, se deberá utilizar el catálogo referencial de valores de mantenimiento correctivo, para el modelo de camionetas adquiridas, otorgado por el proveedor del servicio AMBACAR CIA. LTDA.",
      "El mantenimiento correctivo es estimado, ya que su concurrencia no puede ser planificada. Los costos de repuestos y trabajos que se ejecutaren pueden variar de acuerdo con daños ocultos, los mismos que puede aumentar o disminuir los rubros relacionados a repuestos y mano de obra.",
      "En los casos que exista un desperfecto no contemplado para los trabajos de mantenimiento correctivo, el taller deberá contar con todos los equipos necesarios a fin de atender el requerimiento de reparación de manera ágil y oportuna evitando de esta manera que el vehículo deba permanecer inoperativo por un periodo de tiempo prolongado, excepto en los casos que la rehabilitación amerite."
    ].map(txt => p(txt)),
    
    p("Adicionalmente, se deberá considerar los siguientes aspectos:"),
    ...[
      ["a) ", "Entiéndase como trabajos correctivos aquellos que por ser de carácter imprevisible técnicamente no pueden contemplarse en el Plan de Mantenimiento Preventivo y/o garantía técnica otorgada por el fabricante."],
      ["b) ", "El proveedor tendrá la responsabilidad de realizar la corrección y reparación de averías o fallas mecánicas o de cualquier índole producida de manera espontánea en los vehículos, en base a la orden de mantenimiento."],
      ["c) ", "El proveedor deberá contar con un historial de mantenimientos de la flota vehicular."],
      ["d) ", "El proveedor deberá garantizar el funcionamiento de los automotores una vez salidos del taller, en caso de presentarse inconvenientes con los trabajos mecánicos realizados, se reingresará el vehículo y se solicitará una nueva revisión sin que esta genere costos adicionales."],
      ["e) ", "El proveedor deberá garantizar el stock suficiente, los materiales, repuestos, aditivos o accesorios automotrices, los mismos deben de cumplir con las especificaciones establecidas por el fabricante, con la finalidad de asegurar el cumplimiento de la garantía técnica y el principio de vigencia tecnológica."],
      ["f) ", "El proveedor deberá garantizar el cumplimiento de la garantía técnica en cumplimiento del principio de vigencia tecnológica establecido en la normativa legal vigente, en todos los trabajos desarrollados por el contratista."]
    ].map(([letra, texto]) => p("", {
      leftIndent: 0.5,
      after: 2,
      runs: [
        { text: letra, bold: true },
        { text: texto }
      ]
    })),

    new Paragraph({ spacing: { after: pt(4) } }),

    // Notes (Italic blocks)
    ...[
      ["NOTA: ", "El detalle de mantenimientos antes referido ha sido realizado con base a la información proporcionada por el fabricante, y podrá ser modificado considerando las condiciones operativas y necesidades de la flota vehicular."],
      ["", "Es decir, en el caso de que en el mantenimiento correctivo se consideren varios componentes o ítems que no se encuentren establecidos en los términos de referencia y/o contrato, el contratista deberá emitir un informe técnico y proforma en donde se evidencie que el mantenimiento es conveniente para los intereses institucionales y a su vez mantener la operatividad del vehículo, sin perder la garantía técnica en cumplimiento del principio de vigencia tecnológica."],
      ["", "Por necesidad institucional, los trabajos o el número de veces de ejecución de los ítems indicados podrían variar. Es importante indicar que dichos ítems podrán ser aplicados a una o a todas las camionetas adquiridas a través del Catálogo Electrónico."],
      ["", "Los costos de repuestos y lubricantes son referenciales, y pueden estar sujetos a cambios, debido condiciones de mercado y/o a medidas económicas que afecten directamente a su distribución."]
    ].map(([prefix, txt]) => p("", {
      italic: true,
      after: 3,
      runs: [
        { text: prefix, bold: true, italic: true },
        { text: txt }
      ]
    })),

    new Paragraph({ spacing: { after: pt(6) } }),

    // Consolidado de Valores Table
    heading("CONSOLIDADO DE VALORES"),
    
    (() => {
      const descCorrectivo = incluirCorrectivos
        ? "Se considera un valor aproximado del 30% del costo total de mantenimientos preventivos"
        : "";
      
      const colWidthsCons = [cm(1), cm(4), cm(8), cm(4)];
      
      const headerRow = new TableRow({
        children: [
          cell("", {
            columnSpan: 4,
            bg: "CC0000",
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: pt(2), after: pt(0) },
                children: [
                  new TextRun({
                    text: `COSTOS MANTENIMIENTO ${hojaElegida.toUpperCase()}`,
                    bold: true,
                    size: 18,
                    color: "FFFFFF",
                    font: "Arial"
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: pt(0), after: pt(2) },
                children: [
                  new TextRun({
                    text: `PARA ${cantidadVehiculos} VEHÍCULOS`,
                    bold: true,
                    size: 18,
                    color: "FFFFFF",
                    font: "Arial"
                  })
                ]
              })
            ]
          })
        ]
      });

      const repuestosRow = new TableRow({
        children: [
          cell("1", { align: AlignmentType.CENTER, widthCm: 1 }),
          cell("Total Repuestos", { widthCm: 4 }),
          cell("Mantenimiento preventivo", { rowSpan: 3, widthCm: 8 }),
          cell(`$ ${fmt(totalRepuestos)}`, { align: AlignmentType.RIGHT, widthCm: 4 })
        ]
      });

      const lubricantesRow = new TableRow({
        children: [
          cell("2", { align: AlignmentType.CENTER, widthCm: 1 }),
          cell("Total Lubricantes", { widthCm: 4 }),
          cell(`$ ${fmt(totalLubricantes)}`, { align: AlignmentType.RIGHT, widthCm: 4 })
        ]
      });

      const manoObraRow = new TableRow({
        children: [
          cell("3", { align: AlignmentType.CENTER, widthCm: 1 }),
          cell("Total Mano de Obra", { widthCm: 4 }),
          cell(`$ ${fmt(totalManoObra)}`, { align: AlignmentType.RIGHT, widthCm: 4 })
        ]
      });

      const rows = [headerRow, repuestosRow, lubricantesRow, manoObraRow];

      if (incluirCorrectivos) {
        rows.push(new TableRow({
          children: [
            cell("4", { align: AlignmentType.CENTER, widthCm: 1 }),
            cell("Mantenimiento correctivo", { widthCm: 4 }),
            cell(descCorrectivo, { widthCm: 8 }),
            cell(`$ ${fmt(totalCorrectivo)}`, { align: AlignmentType.RIGHT, widthCm: 4 })
          ]
        }));
      }

      const totalLabel = incluirCorrectivos ? "Total 1+2+3+4" : "Total 1+2+3";
      rows.push(new TableRow({
        children: [
          cell("", { widthCm: 1 }),
          cell("", { widthCm: 4 }),
          cell(totalLabel, { bold: true, align: AlignmentType.RIGHT, bg: "FFCCCC", widthCm: 8 }),
          cell(`$ ${fmt(granTotal)}`, { bold: true, align: AlignmentType.RIGHT, bg: "FFCCCC", widthCm: 4 })
        ]
      }));

      return new Table({
        columnWidths: colWidthsCons,
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
          left: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
          right: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" },
          insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "E0E0E0" }
        },
        rows
      });
    })(),

    p("Costos no incluyen IVA", { align: AlignmentType.CENTER, italic: true, size: 16, before: 4, after: 8 }),

    // Observaciones
    heading("OBSERVACIONES:"),
    ...[
      `Plazo de ejecución de los servicios: ${plazoEjecucion || 365} días calendario, contados a partir del día siguiente de la fecha de suscripción de la orden de compra o contrato.`,
      "Lugar de ejecución del servicio: en las instalaciones de AMBACAR CIA LTDA, a nivel nacional.",
      "Forma de pago:"
    ].map(txt => p("", {
      leftIndent: 0.5,
      after: 3,
      runs: [
        { text: "- ", bold: false },
        { text: txt }
      ]
    })),
    
    // Forma de pago sub-bullets
    ...[
      "100% contra prestación de los servicios parcial, objeto del contrato, de manera mensual.",
      `Mantenimientos preventivos y/o correctivos: se cancelarán una vez que los servicios sean recibidos a satisfacción por parte de la ${razonSocial || 'ENTIDAD CONTRATANTE'}, se haya rendido la garantía técnica y suscrito el acta de entrega recepción parcial y/o definitiva según corresponda.`
    ].map(txt => p("", {
      leftIndent: 1.0,
      after: 3,
      runs: [
        { text: "o  ", bold: false },
        { text: txt }
      ]
    })),
    
    // Garantia Tecnica bullet (has custom formatting)
    p("", {
      leftIndent: 0.5,
      after: 0,
      runs: [
        { text: "- ", bold: false },
        { text: "Garantía Técnica: ", bold: true },
        { text: "Se otorgará una garantía técnica que cubra la calidad de los servicios de mantenimiento preventivo y correctivo, así como de los repuestos utilizados. Dicha garantía tendrá una vigencia de un (1) año para los repuestos, cubriendo defectos de fabricación, y de seis (6) meses para la mano de obra correspondiente a los trabajos ejecutados." }
      ]
    }),
    
    p("La garantía deberá ser presentada al momento de la suscripción de la Orden de Compra o del Contrato, y permanecerá vigente hasta la finalización del último mantenimiento realizado.", {
      leftIndent: 0.8,
      after: 3
    }),

    p("", {
      leftIndent: 0.5,
      after: 3,
      runs: [
        { text: "- ", bold: false },
        { text: `Vigencia de la Proforma: ${vigenciaOferta || '90 días'} calendario, contados a partir de su emisión.` }
      ]
    }),

    new Paragraph({ spacing: { after: pt(8) } }),

    // Firmas
    p("Atentamente,", { size: 18, after: 30 }),
    p("Ing. Luis Vintimilla", { size: 18, after: 0 }),
    p("APODERADO ESPECIAL", { bold: true, size: 18, after: 0 }),
    p("AMBACAR CIA. LTDA.", { bold: true, size: 18, after: 0 })
  ];

  // --- CREATE THE DOCUMENT STRUCTURE ---
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: {
            width: 11907, // A4
            height: 16840,
          },
          margin: {
            top: 1134, // 2cm
            right: 1134,
            bottom: 1134,
            left: 1134,
            header: 567, // 1cm
            footer: 680 // ~1.2cm
          }
        }
      },
      headers: {
        default: new Header({
          children: headerParagraphs
        })
      },
      footers: {
        default: new Footer({
          children: footerParagraphs
        })
      },
      children: docChildren
    }]
  });

  // Pack the document and trigger download
  const blob = await Packer.toBlob(doc);
  const nombreDocx = `Proforma_${numProforma}.docx`;
  
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombreDocx;
  a.click();
  window.URL.revokeObjectURL(url);
}
