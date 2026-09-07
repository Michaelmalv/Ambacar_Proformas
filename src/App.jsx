import React, { useState, useEffect } from 'react';
import { 
  UploadCloud, 
  FileSpreadsheet, 
  FileCode, 
  CheckCircle2, 
  AlertTriangle, 
  Trash2, 
  Plus, 
  Download, 
  Image as ImageIcon,
  Building2, 
  Hash, 
  User, 
  Phone, 
  Mail, 
  MapPin, 
  Calendar,
  Database,
  History,
  Car,
  Layers,
  FileText,
  Sliders,
  X
} from 'lucide-react';
import { 
  buscarValor, 
  parseRangoKM, 
  detectarCorrectivosExcel, 
  parsePlacas, 
  parseMantenimientosObservacion,
  detectarModeloEnGrid,
  calcularPlanMantenimiento,
  roundMoney
} from './utils/excelParser';
import { generarProformaDocx } from './utils/docxGenerator';
import { 
  DEFAULT_MODELOS,
  getModelosVehiculo, 
  buscarPlanPorModelo, 
  getCostosPorKm, 
  guardarProforma, 
  getHistorialProformas 
} from './services/supabase';

export default function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState('generator'); // 'generator' | 'history'

  // Files State & Drag States
  const [cotizacionFile, setCotizacionFile] = useState(null);
  const [isDraggingCot, setIsDraggingCot] = useState(false);
  const [templateFile, setTemplateFile] = useState(null);
  const [isDraggingTemplate, setIsDraggingTemplate] = useState(false);
  const [isDraggingHeader, setIsDraggingHeader] = useState(false);
  const [isDraggingFooter, setIsDraggingFooter] = useState(false);
  const [showManualTemplate, setShowManualTemplate] = useState(false);
  
  // Supabase State (Pre-populated synchronously so dropdown is never blank)
  const [modelosDisponibles, setModelosDisponibles] = useState(DEFAULT_MODELOS);
  const [selectedPlanId, setSelectedPlanId] = useState(DEFAULT_MODELOS[0]?.planes_mantenimiento?.[0]?.id || "w7-4x4-250k");
  const [selectedModeloId, setSelectedModeloId] = useState(DEFAULT_MODELOS[0]?.id || "w7-4x4-250k");
  const [supabaseConnected, setSupabaseConnected] = useState(false);
  const [historialProformas, setHistorialProformas] = useState([]);
  const [loadingHistory, setLoadingHistory] = useState(false);

  // Maintenance Calculation Mode State: 'individual' (per vehicle from observation) vs 'rango' (global range)
  const [modoCalculo, setModoCalculo] = useState('auto'); // 'auto' | 'individual' | 'rango'
  const [mantenimientosEspecificos, setMantenimientosEspecificos] = useState([]);
  const [newVehiculoPlaca, setNewVehiculoPlaca] = useState("");
  const [newVehiculoKm, setNewVehiculoKm] = useState("35000");

  // Manual Template Workbooks & sheets state
  const [templateWorkbook, setTemplateWorkbook] = useState(null);
  const [hojas, setHojas] = useState([]);
  const [sheetSelection, setSheetSelection] = useState("");

  // Form Fields State
  const [form, setForm] = useState({
    razonSocial: "",
    ruc: "",
    direccion: "",
    contacto: "",
    telefono: "",
    correo: "",
    modeloVehiculo: DEFAULT_MODELOS[0]?.nombre_completo || "WINGLE 7 DIESEL 4X4",
    cantidadVehiculos: 1,
    plazoEjecucion: "365",
    vigenciaOferta: "90 días",
    observacion: "",
    objetoContrato: "Plan de mantenimiento preventivo y correctivo de vehículos",
    kmDesde: "5000",
    kmHasta: "120000",
    incluirCorrectivos: false
  });
  
  const [placas, setPlacas] = useState([]);
  const [newPlate, setNewPlate] = useState("");

  // Base64 Images for docx
  const [headerImg, setHeaderImg] = useState(null);
  const [footerImg, setFooterImg] = useState(null);

  // Calculation Results
  const [calcResult, setCalcResult] = useState(null);
  const [calcError, setCalcError] = useState(null);
  const [isGenerating, setIsGenerating] = useState(false);

  // Load Models from Supabase on mount
  useEffect(() => {
    async function loadSupabaseData() {
      const { data, error, source } = await getModelosVehiculo();
      if (!error && data && data.length > 0) {
        setModelosDisponibles(data);
        if (source === 'supabase') {
          setSupabaseConnected(true);
          setSelectedModeloId(prev => {
            const match = data.find(m => m.id === prev || m.nombre_completo === DEFAULT_MODELOS[0]?.nombre_completo);
            return match ? match.id : data[0].id;
          });
          setSelectedPlanId(prev => {
            const match = data.find(m => m.id === selectedModeloId || m.nombre_completo === DEFAULT_MODELOS[0]?.nombre_completo);
            return (match && match.planes_mantenimiento?.[0]?.id) || data[0]?.planes_mantenimiento?.[0]?.id || prev;
          });
        }
      }
    }
    loadSupabaseData();
  }, []);

  // Load cached header/footer images from LocalStorage on mount
  useEffect(() => {
    const loadDefaultHeader = async () => {
      const cachedHeader = localStorage.getItem("ambacar_header_img");
      if (cachedHeader) {
        setHeaderImg(cachedHeader);
      } else {
        try {
          const response = await fetch("/header_ambacar.png");
          if (response.ok) {
            const blob = await response.blob();
            const reader = new FileReader();
            reader.onloadend = () => {
              setHeaderImg(reader.result);
            };
            reader.readAsDataURL(blob);
          }
        } catch (err) {
          console.error("Error loading default header image:", err);
        }
      }
    };
    loadDefaultHeader();

    const cachedFooter = localStorage.getItem("ambacar_footer_img");
    if (cachedFooter) setFooterImg(cachedFooter);
  }, []);

  // Load Proformas history when switching tab
  useEffect(() => {
    if (activeTab === 'history') {
      setLoadingHistory(true);
      getHistorialProformas().then(({ data }) => {
        setHistorialProformas(data || []);
        setLoadingHistory(false);
      });
    }
  }, [activeTab]);

  // Global window drag & drop prevention to ensure files are never downloaded/opened by the browser
  useEffect(() => {
    const preventDefaults = (e) => {
      e.preventDefault();
      e.stopPropagation();
    };
    window.addEventListener('dragover', preventDefaults, false);
    window.addEventListener('drop', preventDefaults, false);
    return () => {
      window.removeEventListener('dragover', preventDefaults);
      window.removeEventListener('drop', preventDefaults);
    };
  }, []);

  // Process Quotation File (DATOS_PARA_COTIZACION.xlsx)
  const processCotizacionFile = (file) => {
    if (!file) return;
    setCotizacionFile(file);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const importXLSX = await import('xlsx');
        const workbook = importXLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const grid = importXLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        // Extract fields with multi-label aliases
        const razonSocialVal = buscarValor(grid, ["Razón Social", "Razon Social", "Cliente", "Nombre del Cliente", "Empresa"]) || "";
        const rucVal = buscarValor(grid, ["RUC", "R.U.C.", "C.I.", "Identificación", "Identificacion"]) || "";
        const direccionVal = buscarValor(grid, ["Dirección", "Direccion", "Ubicación", "Ubicacion", "Domicilio"]) || "";
        const contactoVal = buscarValor(grid, ["Nombre del Contacto", "Contacto", "Atención", "Atencion", "Solicitante", "Persona de Contacto"]) || "";
        const telefonoVal = buscarValor(grid, ["Teléfono", "Telefono", "Celular", "Telf", "Movil"]) || "";
        const correoVal = buscarValor(grid, ["Correo electrónico", "Correo electronico", "Correo", "Email", "E-mail"]) || "";
        const plazoVal = buscarValor(grid, ["Plazo de ejecución", "Plazo de ejecucion", "Plazo"]) || "365";
        const vigenciaVal = buscarValor(grid, ["Vigencia de la oferta", "Vigencia"]) || "90 días";
        const modeloVal = buscarValor(grid, [
          "Modelo de vehículo", "Modelo de vehiculo", "Modelo del vehículo", "Modelo del vehiculo", 
          "Modelo", "Vehículo", "Vehiculo", "Vehículos", "Vehiculos", "Tipo de Vehículo", "Tipo de vehiculo", 
          "Flota", "Camioneta", "Jeep", "Automotor", "Unidad", "Descripción", "Descripcion", "Detalle", 
          "Bien / Servicio", "Producto", "Item"
        ]) || "";
        const gridModel = detectarModeloEnGrid(grid);
        const finalModeloVal = modeloVal || gridModel || "";

        const cantidadVal = buscarValor(grid, [
          "Cantidad de vehículos", "Cantidad de vehiculos", "Cantidad", "N° vehículos", "N° vehiculos", 
          "No. Vehículos", "No. Vehiculos", "Cant.", "Cant", "Total Vehículos", "Unidades"
        ]) || "1";
        const observacionVal = buscarValor(grid, ["Observación", "Observacion", "Notas", "Observaciones"]) || "";
        const objetoVal = buscarValor(grid, ["OBJETO DEL CONTRATO", "Objeto del contrato", "Objeto de la contratación", "Objeto"]) || "Plan de mantenimiento preventivo y correctivo de vehículos";
        
        const { kmDesde: kd, kmHasta: kh } = parseRangoKM(grid);
        const placasVal = parsePlacas(grid);
        
        // Detect checkboxes using zip vml
        let inclCorr = false;
        try {
          const zipResult = await detectarCorrectivosExcel(data, grid);
          if (zipResult) {
            inclCorr = zipResult.siChecked && !zipResult.noChecked;
          } else {
            // Fallback text check
            for (let r = 0; r < grid.length; r++) {
              const row = grid[r] || [];
              for (let c = 0; c < row.length; c++) {
                const cellText = (row[c] || '').toString().toLowerCase();
                if (cellText.includes("correctivo") && cellText.includes("incluir")) {
                  for (let vc = c + 1; vc < row.length; vc++) {
                    const v = (row[vc] || '').toString().trim().toUpperCase();
                    if (v === "SI" || v === "SÍ") inclCorr = true;
                  }
                }
              }
            }
          }
        } catch (err) {
          console.warn("Checkbox detection error:", err);
        }

        // Parse individual vehicle maintenances from observation
        const parsedMant = parseMantenimientosObservacion(observacionVal);
        if (parsedMant && parsedMant.length > 0) {
          setMantenimientosEspecificos(parsedMant);
          setModoCalculo('individual');
        } else {
          setMantenimientosEspecificos([]);
          setModoCalculo('rango');
        }

        // Match with Supabase / Fallback Model & Plan (considering model text, grid scan, and observation)
        const { plan, modelo } = await buscarPlanPorModelo(finalModeloVal, observacionVal);
        if (plan && modelo) {
          const activeModel = modelosDisponibles.find(m => 
            m.id === modelo.id || 
            m.nombre_completo?.toUpperCase() === modelo.nombre_completo?.toUpperCase() ||
            m.modelo?.toUpperCase() === modelo.modelo?.toUpperCase()
          );

          const finalModId = activeModel ? activeModel.id : modelo.id;
          const finalPlanId = (activeModel && activeModel.planes_mantenimiento?.[0]?.id) || plan.id;

          setSelectedModeloId(finalModId);
          setSelectedPlanId(finalPlanId);
        }

        let parsedCount = parseInt(cantidadVal, 10);
        if (isNaN(parsedCount) || parsedCount <= 0) parsedCount = 1;
        if (placasVal && placasVal.length > parsedCount) {
          parsedCount = placasVal.length;
        }
        const effectiveCantidad = (parsedMant && parsedMant.length > 0) ? parsedMant.length : parsedCount;

        setForm(prev => ({
          ...prev,
          razonSocial: razonSocialVal || prev.razonSocial,
          ruc: rucVal || prev.ruc,
          direccion: direccionVal || prev.direccion,
          contacto: contactoVal || prev.contacto,
          telefono: telefonoVal || prev.telefono,
          correo: correoVal || prev.correo,
          plazoEjecucion: plazoVal || prev.plazoEjecucion,
          vigenciaOferta: vigenciaVal || prev.vigenciaOferta,
          modeloVehiculo: finalModeloVal || modelo?.nombre_completo || prev.modeloVehiculo,
          cantidadVehiculos: effectiveCantidad,
          observacion: observacionVal || prev.observacion,
          objetoContrato: objetoVal || prev.objetoContrato,
          kmDesde: kd ? kd.toString() : (prev.kmDesde || "5000"),
          kmHasta: kh ? kh.toString() : (prev.kmHasta || "120000"),
          incluirCorrectivos: inclCorr
        }));
        
        if (parsedMant && parsedMant.length > 0) {
          setPlacas(parsedMant.map(m => m.placa));
        } else if (placasVal && placasVal.length > 0) {
          setPlacas(placasVal);
        }
      } catch (err) {
        console.error("Error reading quotation file:", err);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleCotizacionChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      processCotizacionFile(e.target.files[0]);
    }
  };

  // Process Maintenance Plan Template File (Offline Manual Fallback)
  const processTemplateFile = (file) => {
    if (!file) return;
    setTemplateFile(file);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const importXLSX = await import('xlsx');
        const workbook = importXLSX.read(data, { type: 'array' });
        setTemplateWorkbook(workbook);
        setHojas(workbook.SheetNames);
        
        const modelUpper = (form.modeloVehiculo || "").toUpperCase();
        let detected = null;
        const exclude = ["CORRECTIVO", "CONSOLID", "CRONOGRAMA"];
        
        for (const name of workbook.SheetNames) {
          if (exclude.some(ex => name.toUpperCase().includes(ex))) continue;
          if (modelUpper.includes("4X4") && name.toUpperCase().includes("4X4")) {
            detected = name;
            break;
          }
          if (modelUpper.includes("4X2") && name.toUpperCase().includes("4X2")) {
            detected = name;
            break;
          }
        }
        
        setSheetSelection(detected || workbook.SheetNames[0] || "");
      } catch (err) {
        console.error("Error reading template file:", err);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleTemplateChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      processTemplateFile(e.target.files[0]);
    }
  };

  // Handle Model Selection change in dropdown
  const handleModeloSelectChange = (e) => {
    const modeloId = e.target.value;
    setSelectedModeloId(modeloId);
    const selectedMod = modelosDisponibles.find(m => m.id === modeloId);
    if (selectedMod && selectedMod.planes_mantenimiento && selectedMod.planes_mantenimiento.length > 0) {
      setSelectedPlanId(selectedMod.planes_mantenimiento[0].id);
      setForm(prev => ({
        ...prev,
        modeloVehiculo: selectedMod.nombre_completo
      }));
    }
  };

  // Handle Observation Text Changes (allows live re-parsing of specific maintenances)
  const handleObservacionChange = (e) => {
    const newObs = e.target.value;
    setForm(prev => ({ ...prev, observacion: newObs }));
    const parsed = parseMantenimientosObservacion(newObs);
    if (parsed.length > 0) {
      setMantenimientosEspecificos(parsed);
      setForm(prev => ({ ...prev, cantidadVehiculos: parsed.length }));
      setPlacas(parsed.map(p => p.placa));
    }
  };

  // Add a vehicle to individual maintenance list
  const handleAddVehiculoMant = (e) => {
    e.preventDefault();
    const kmNum = parseInt(newVehiculoKm, 10);
    if (isNaN(kmNum) || kmNum < 5000) return;
    
    const placaName = newVehiculoPlaca.trim().toUpperCase() || `Vehículo ${mantenimientosEspecificos.length + 1}`;
    const newItem = {
      id: `item-${Date.now()}`,
      placa: placaName,
      km: kmNum,
      raw: `${placaName} = MANTTO DE ${kmNum.toLocaleString()} KM`
    };

    const updated = [...mantenimientosEspecificos, newItem];
    setMantenimientosEspecificos(updated);
    setForm(prev => ({ ...prev, cantidadVehiculos: updated.length }));
    setPlacas(updated.map(p => p.placa));
    setNewVehiculoPlaca("");
  };

  // Remove a vehicle from individual maintenance list
  const handleRemoveVehiculoMant = (idx) => {
    const updated = mantenimientosEspecificos.filter((_, i) => i !== idx);
    setMantenimientosEspecificos(updated);
    setForm(prev => ({ ...prev, cantidadVehiculos: Math.max(1, updated.length) }));
    setPlacas(updated.map(p => p.placa));
  };

  // Recalculate cost when range, vehicles, plan, or mode changes
  useEffect(() => {
    const kDesde = parseInt(form.kmDesde, 10);
    const kHasta = parseInt(form.kmHasta, 10);
    const nVehic = parseInt(form.cantidadVehiculos, 10) || 1;

    const isIndividualMode = (modoCalculo === 'individual' || (modoCalculo === 'auto' && mantenimientosEspecificos.length > 0)) && mantenimientosEspecificos.length > 0;

    // 1. If using Supabase / Bundled Plan (Default Mode)
    if (selectedPlanId && !templateWorkbook) {
      if (isIndividualMode) {
        getCostosPorKm(selectedPlanId, null, null, mantenimientosEspecificos.length, mantenimientosEspecificos).then(res => {
          if (res.error) {
            setCalcResult(null);
            setCalcError(res.message || "Error al calcular desde base de datos");
          } else {
            setCalcResult(res);
            setCalcError(null);
          }
        });
      } else {
        if (isNaN(kDesde) || isNaN(kHasta) || kDesde <= 0 || kHasta <= 0) {
          return;
        }
        getCostosPorKm(selectedPlanId, kDesde, kHasta, nVehic, null).then(res => {
          if (res.error) {
            setCalcResult(null);
            setCalcError(res.message || "Error al calcular desde base de datos");
          } else {
            setCalcResult(res);
            setCalcError(null);
          }
        });
      }
      return;
    }

    // 2. If using Manual Template Excel (Fallback Mode)
    if (templateWorkbook && sheetSelection) {
      try {
        const worksheet = templateWorkbook.Sheets[sheetSelection];
        if (!worksheet) return;
        
        import('xlsx').then(XLSXLib => {
          const planGrid = XLSXLib.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          const res = calcularPlanMantenimiento(planGrid, kDesde, kHasta, nVehic);
          if (res.error) {
            setCalcResult(null);
            setCalcError(`No se encontraron kilómetros válidos en el rango [${kDesde} - ${kHasta} KM]. KMs en plantilla: ${res.kmsDisponibles.join(', ')}`);
          } else {
            setCalcResult(res);
            setCalcError(null);
          }
        });
      } catch (err) {
        console.error("Error manual template calculation:", err);
        setCalcError("Error al procesar plantilla manual.");
      }
    }
  }, [selectedPlanId, templateWorkbook, sheetSelection, form.kmDesde, form.kmHasta, form.cantidadVehiculos, modoCalculo, mantenimientosEspecificos]);

  // Form input changes
  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setForm(prev => ({
      ...prev,
      [name]: type === 'checkbox' ? checked : value
    }));
  };

  // Handle image uploads
  const handleImageUpload = (file, type) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const base64 = e.target.result;
      if (type === 'header') {
        setHeaderImg(base64);
        localStorage.setItem("ambacar_header_img", base64);
      } else {
        setFooterImg(base64);
        localStorage.setItem("ambacar_footer_img", base64);
      }
    };
    reader.readAsDataURL(file);
  };

  const clearCachedImage = (type) => {
    if (type === 'header') {
      setHeaderImg(null);
      localStorage.removeItem("ambacar_header_img");
    } else {
      setFooterImg(null);
      localStorage.removeItem("ambacar_footer_img");
    }
  };

  // Add Plate
  const handleAddPlate = (e) => {
    e.preventDefault();
    if (newPlate.trim()) {
      setPlacas([...placas, newPlate.trim().toUpperCase()]);
      setNewPlate("");
    }
  };

  const handleRemovePlate = (index) => {
    setPlacas(placas.filter((_, idx) => idx !== index));
  };

  // Trigger Word docx Generation and save to Supabase
  const handleGenerateWord = async () => {
    if (!calcResult) return;
    setIsGenerating(true);
    
    const totalRepuestos = roundMoney(calcResult.totalRepuestos);
    const totalLubricantes = roundMoney(calcResult.totalLubricantes);
    const totalManoObra = roundMoney(calcResult.totalMo);
    const totalPreventivo = roundMoney(calcResult.totalPreventivo);
    const totalCorrectivo = form.incluirCorrectivos ? roundMoney(totalPreventivo * 0.3) : 0;
    const granTotal = roundMoney(totalPreventivo + totalCorrectivo);

    const today = new Date();
    const mmStr = String(today.getMonth() + 1).padStart(2, '0');
    const ddStr = String(today.getDate()).padStart(2, '0');
    const numProforma = `AMB-PV-${today.getFullYear()}-AUTO-${mmStr}${ddStr}`;

    const currentModelName = modelosDisponibles.find(m => m.id === selectedModeloId)?.nombre_completo || form.modeloVehiculo;

    const documentData = {
      ...form,
      numeroProforma: numProforma,
      placas,
      totalRepuestos,
      totalLubricantes,
      totalManoObra,
      totalPreventivo,
      totalCorrectivo,
      granTotal,
      hojaElegida: currentModelName || sheetSelection || "MANTENIMIENTO",
      selectedPlanId: selectedPlanId,
      headerImage: headerImg,
      footerImage: footerImg,
      tablaDetalle: calcResult.tablaDetalle
    };

    try {
      // 1. Generate & Download Word Document (.docx)
      await generarProformaDocx(documentData);

      // 2. Save Proforma Record in Supabase
      if (supabaseConnected) {
        await guardarProforma(documentData);
      }
    } catch (err) {
      console.error("Error generating docx or saving proforma:", err);
      alert("Error al generar la proforma.");
    } finally {
      setIsGenerating(false);
    }
  };

  const getFmtCurrency = (val) => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD'
    }).format(val || 0);
  };

  const previewPrev = calcResult ? roundMoney(calcResult.totalPreventivo) : 0;
  const previewCorr = form.incluirCorrectivos ? roundMoney(previewPrev * 0.3) : 0;
  const previewTotal = roundMoney(previewPrev + previewCorr);

  return (
    <div className="app-container">
      {/* Header section */}
      <header className="app-header">
        <div className="brand-section">
          <img src="/ambacar_logo.png" alt="Ambacar" className="brand-logo-img" />
          <div className="brand-divider"></div>
          <div>
            <h1 className="brand-title"><span>AMBACAR</span> PROFORMAS</h1>
            <p className="brand-subtitle">Cotizaciones Corporativas conectadas con Supabase</p>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          {supabaseConnected && (
            <div className="supabase-badge">
              <span className="status-dot"></span>
              <span>Supabase Conectado ({modelosDisponibles.length} modelos)</span>
            </div>
          )}
          <div className="brand-badge">CLOUD V2.0</div>
        </div>
      </header>

      {/* Navigation Tabs */}
      <div className="tab-navigation">
        <button 
          className={`tab-btn ${activeTab === 'generator' ? 'active' : ''}`}
          onClick={() => setActiveTab('generator')}
        >
          <Car size={16} />
          Generador de Proforma
        </button>
        <button 
          className={`tab-btn ${activeTab === 'history' ? 'active' : ''}`}
          onClick={() => setActiveTab('history')}
        >
          <History size={16} />
          Historial en Supabase
        </button>
      </div>

      {activeTab === 'generator' ? (
        /* Grid container for Generator */
        <div className="dashboard-grid">
          
          {/* Left Column: Form & Config */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
            
            {/* Card 1: File Uploader */}
            <div className="card">
              <h3 className="card-title">
                <FileSpreadsheet size={20} className="text-muted" />
                1. Cargar Cotización del Cliente
              </h3>
              
              <div className="form-group" style={{ marginBottom: '1rem' }}>
                <label className="form-label">Archivo de Cotización (.xlsx)</label>
                {!cotizacionFile ? (
                  <div 
                    className={`upload-zone ${isDraggingCot ? 'dragging' : ''}`} 
                    onClick={() => document.getElementById('cot-input').click()}
                    onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingCot(true); }}
                    onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingCot(true); }}
                    onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingCot(false); }}
                    onDrop={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      setIsDraggingCot(false);
                      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                        processCotizacionFile(e.dataTransfer.files[0]);
                      }
                    }}
                  >
                    <UploadCloud className="upload-icon" />
                    <p className="upload-text">Arrastra o selecciona el archivo</p>
                    <p className="upload-subtext">DATOS_PARA_COTIZACION.xlsx</p>
                    <input 
                      type="file" 
                      id="cot-input" 
                      style={{ display: 'none' }} 
                      accept=".xlsx" 
                      onChange={handleCotizacionChange}
                    />
                  </div>
                ) : (
                  <div className="file-pill">
                    <div className="file-pill-info">
                      <CheckCircle2 size={16} style={{ color: 'var(--success)' }} />
                      <span>{cotizacionFile.name}</span>
                    </div>
                    <button className="file-pill-remove" onClick={() => { setCotizacionFile(null); setCalcResult(null); }}>
                      <X size={14} />
                    </button>
                  </div>
                )}
              </div>

              {/* Model selection from Supabase */}
              <div className="form-group" style={{ marginBottom: 0 }}>
                <label className="form-label">
                  <Database size={14} style={{ display: 'inline', marginRight: '4px', verticalAlign: 'text-bottom' }} />
                  Plan de Mantenimiento (Base de Datos):
                </label>
                <select 
                  className="form-input"
                  value={selectedModeloId}
                  onChange={handleModeloSelectChange}
                >
                  {modelosDisponibles.map(m => (
                    <option key={m.id} value={m.id}>
                      {m.nombre_completo}
                    </option>
                  ))}
                </select>
                <span className="upload-subtext" style={{ marginTop: '0.25rem', display: 'block' }}>
                  ⚡ Se detecta automáticamente del Excel o puedes cambiarlo aquí.
                </span>
              </div>

              {/* Optional Manual Template Excel Accordion */}
              <div style={{ marginTop: '1rem', borderTop: '1px solid var(--border-light)', paddingTop: '0.75rem' }}>
                <button 
                  type="button" 
                  onClick={() => setShowManualTemplate(!showManualTemplate)}
                  style={{ background: 'none', border: 'none', color: 'var(--text-muted)', fontSize: '0.75rem', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '0.35rem' }}
                >
                  <Layers size={14} />
                  {showManualTemplate ? 'Ocultar carga manual de plantilla Excel' : '¿Deseas usar una plantilla Excel local en lugar de Supabase?'}
                </button>

                {showManualTemplate && (
                  <div style={{ marginTop: '0.75rem' }}>
                    {!templateFile ? (
                      <div 
                        className={`upload-zone ${isDraggingTemplate ? 'dragging' : ''}`} 
                        style={{ padding: '1rem' }} 
                        onClick={() => document.getElementById('template-input').click()}
                        onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(true); }}
                        onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(true); }}
                        onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(false); }}
                        onDrop={(e) => {
                          e.preventDefault();
                          e.stopPropagation();
                          setIsDraggingTemplate(false);
                          if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                            processTemplateFile(e.dataTransfer.files[0]);
                          }
                        }}
                      >
                        <UploadCloud size={20} style={{ margin: '0 auto 0.25rem' }} />
                        <span className="upload-subtext">Subir Plantilla Excel Manual</span>
                        <input 
                          type="file" 
                          id="template-input" 
                          style={{ display: 'none' }} 
                          accept=".xlsx" 
                          onChange={handleTemplateChange}
                        />
                      </div>
                    ) : (
                      <div className="file-pill">
                        <div className="file-pill-info">
                          <CheckCircle2 size={16} style={{ color: 'var(--success)' }} />
                          <span>{templateFile.name}</span>
                        </div>
                        <button className="file-pill-remove" onClick={() => { setTemplateFile(null); setTemplateWorkbook(null); setHojas([]); setSheetSelection(""); }}>
                          <X size={14} />
                        </button>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>

            {/* Card 2: Header / Footer images */}
            <div className="card">
              <h3 className="card-title">
                <ImageIcon size={20} className="text-muted" />
                Cabecera y Pie de Página
              </h3>
              
              <div className="image-uploader-container">
                {/* Header Box */}
                <div>
                  <label className="form-label">Cabecera (Header Ambacar)</label>
                  {headerImg ? (
                    <div className="image-preview-box">
                      <img src={headerImg} alt="Header Preview" />
                      <button 
                        className="file-pill-remove" 
                        style={{ position: 'absolute', top: '5px', right: '5px', backgroundColor: 'rgba(0,0,0,0.5)' }} 
                        onClick={() => clearCachedImage('header')}
                      >
                        <X size={14} />
                      </button>
                    </div>
                  ) : (
                    <div 
                      className={`upload-zone ${isDraggingHeader ? 'dragging' : ''}`} 
                      style={{ padding: '1rem' }} 
                      onClick={() => document.getElementById('header-img-input').click()}
                      onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingHeader(true); }}
                      onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingHeader(true); }}
                      onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingHeader(false); }}
                      onDrop={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        setIsDraggingHeader(false);
                        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                          handleImageUpload(e.dataTransfer.files[0], 'header');
                        }
                      }}
                    >
                      <UploadCloud size={20} style={{ margin: '0 auto 0.25rem' }} />
                      <span className="upload-subtext">Subir Header</span>
                      <input 
                        type="file" 
                        id="header-img-input" 
                        style={{ display: 'none' }} 
                        accept="image/*" 
                        onChange={(e) => handleImageUpload(e.target.files[0], 'header')}
                      />
                    </div>
                  )}
                </div>

                {/* Footer Box */}
                <div>
                  <label className="form-label">Pie de Página (Footer Data)</label>
                  {footerImg ? (
                    <div className="image-preview-box">
                      <img src={footerImg} alt="Footer Preview" />
                      <button 
                        className="file-pill-remove" 
                        style={{ position: 'absolute', top: '5px', right: '5px', backgroundColor: 'rgba(0,0,0,0.5)' }} 
                        onClick={() => clearCachedImage('footer')}
                      >
                        <X size={14} />
                      </button>
                    </div>
                  ) : (
                    <div 
                      className={`upload-zone ${isDraggingFooter ? 'dragging' : ''}`} 
                      style={{ padding: '1rem' }} 
                      onClick={() => document.getElementById('footer-img-input').click()}
                      onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingFooter(true); }}
                      onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingFooter(true); }}
                      onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingFooter(false); }}
                      onDrop={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        setIsDraggingFooter(false);
                        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                          handleImageUpload(e.dataTransfer.files[0], 'footer');
                        }
                      }}
                    >
                      <UploadCloud size={20} style={{ margin: '0 auto 0.25rem' }} />
                      <span className="upload-subtext">Subir Footer</span>
                      <input 
                        type="file" 
                        id="footer-img-input" 
                        style={{ display: 'none' }} 
                        accept="image/*" 
                        onChange={(e) => handleImageUpload(e.target.files[0], 'footer')}
                      />
                    </div>
                  )}
                </div>
              </div>
              
              {!footerImg && (
                <p className="upload-subtext" style={{ marginTop: '0.75rem', color: '#9ca3af', display: 'flex', gap: '0.25rem', alignItems: 'center' }}>
                  ℹ️ Se incluye texto legal por defecto para el pie de página.
                </p>
              )}
            </div>

            {/* Card 3: Datos del Cliente & Formulario */}
            <div className="card">
              <h3 className="card-title">
                <Building2 size={20} className="text-muted" />
                Datos del Formulario
              </h3>
              
              <div className="form-group">
                <label className="form-label">Objeto del Contrato</label>
                <input 
                  className="form-input" 
                  name="objetoContrato" 
                  value={form.objetoContrato} 
                  onChange={handleChange} 
                />
              </div>

              <div className="form-group half-width">
                <div>
                  <label className="form-label">Razón Social Cliente</label>
                  <input 
                    className="form-input" 
                    name="razonSocial" 
                    value={form.razonSocial} 
                    onChange={handleChange} 
                  />
                </div>
                <div>
                  <label className="form-label">RUC Cliente</label>
                  <input 
                    className="form-input" 
                    name="ruc" 
                    value={form.ruc} 
                    onChange={handleChange} 
                  />
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Dirección Cliente</label>
                <input 
                  className="form-input" 
                  name="direccion" 
                  value={form.direccion} 
                  onChange={handleChange} 
                />
              </div>

              <div className="form-group half-width">
                <div>
                  <label className="form-label">Contacto (Nombre)</label>
                  <input 
                    className="form-input" 
                    name="contacto" 
                    value={form.contacto} 
                    onChange={handleChange} 
                  />
                </div>
                <div>
                  <label className="form-label">Teléfono</label>
                  <input 
                    className="form-input" 
                    name="telefono" 
                    value={form.telefono} 
                    onChange={handleChange} 
                  />
                </div>
              </div>

              <div className="form-group">
                <label className="form-label">Correo Electrónico</label>
                <input 
                  className="form-input" 
                  name="correo" 
                  value={form.correo} 
                  onChange={handleChange} 
                />
              </div>

              <div className="form-group half-width">
                <div>
                  <label className="form-label">Modelo del Vehículo</label>
                  <input 
                    className="form-input" 
                    name="modeloVehiculo" 
                    value={form.modeloVehiculo} 
                    placeholder="Ej: POER 4X2"
                    onChange={handleChange} 
                  />
                </div>
                <div>
                  <label className="form-label">Cantidad Vehículos</label>
                  <input 
                    type="number"
                    className="form-input" 
                    name="cantidadVehiculos" 
                    value={form.cantidadVehiculos} 
                    onChange={handleChange} 
                  />
                </div>
              </div>

              {/* OBSERVACIÓN FIELD */}
              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span>
                    <FileText size={14} style={{ display: 'inline', marginRight: '4px', verticalAlign: 'text-bottom' }} />
                    Observación del Cliente (Mantenimientos Necesarios)
                  </span>
                  {mantenimientosEspecificos.length > 0 && (
                    <span style={{ fontSize: '0.72rem', color: 'var(--success)', fontWeight: 600 }}>
                      ⚡ {mantenimientosEspecificos.length} vehículos detectados
                    </span>
                  )}
                </label>
                <textarea 
                  className="form-input" 
                  rows={3}
                  name="observacion" 
                  value={form.observacion} 
                  placeholder="Ej: HEI-1859= MANTTO DE 35.000 KM&#10;HEI-1845= MANTTO DE 30.000 KM&#10;HEI-1862= MANTTO DE 35.000 KM"
                  onChange={handleObservacionChange}
                  style={{ fontFamily: 'monospace', fontSize: '0.82rem', resize: 'vertical' }}
                />
              </div>

              {/* CALCULATION MODE SWITCHER */}
              <div className="form-group">
                <label className="form-label">
                  <Sliders size={14} style={{ display: 'inline', marginRight: '4px', verticalAlign: 'text-bottom' }} />
                  Modalidad de Cálculo de Mantenimientos
                </label>
                <div style={{ display: 'flex', gap: '0.5rem' }}>
                  <button 
                    type="button"
                    className={`btn ${modoCalculo === 'individual' || (modoCalculo === 'auto' && mantenimientosEspecificos.length > 0) ? 'btn-primary' : 'btn-secondary'}`}
                    style={{ flex: 1, padding: '0.55rem', fontSize: '0.78rem' }}
                    onClick={() => setModoCalculo('individual')}
                  >
                    🚗 Por Vehículo Individual ({mantenimientosEspecificos.length || form.cantidadVehiculos})
                  </button>
                  <button 
                    type="button"
                    className={`btn ${modoCalculo === 'rango' ? 'btn-primary' : 'btn-secondary'}`}
                    style={{ flex: 1, padding: '0.55rem', fontSize: '0.78rem' }}
                    onClick={() => setModoCalculo('rango')}
                  >
                    📊 Rango Global ({form.kmDesde || '5000'}k - {form.kmHasta || '120k'}k)
                  </button>
                </div>
              </div>

              {/* Individual Vehicle Maintenance List (if individual mode active) */}
              {(modoCalculo === 'individual' || (modoCalculo === 'auto' && mantenimientosEspecificos.length > 0)) && (
                <div className="form-group" style={{ background: '#f8fafc', padding: '0.85rem', borderRadius: '8px', border: '1px solid var(--border-light)' }}>
                  <label className="form-label" style={{ marginBottom: '0.5rem', fontSize: '0.78rem' }}>
                    Vehículos y Kilometrajes Asignados:
                  </label>
                  
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '0.4rem', marginBottom: '0.75rem' }}>
                    {mantenimientosEspecificos.map((item, idx) => (
                      <div key={item.id || idx} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#ffffff', border: '1px solid var(--border-light)', padding: '0.45rem 0.65rem', borderRadius: '6px', boxShadow: 'var(--shadow-sm)' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                          <span style={{ fontWeight: 700, color: 'var(--primary)', fontSize: '0.85rem' }}>{item.placa}</span>
                          <span style={{ color: 'var(--text-muted)', fontSize: '0.78rem' }}>➜</span>
                          <span style={{ fontWeight: 600, color: 'var(--text-main)', fontSize: '0.85rem' }}>{item.km.toLocaleString()} KM</span>
                        </div>
                        <button 
                          type="button"
                          onClick={() => handleRemoveVehiculoMant(idx)}
                          style={{ background: 'none', border: 'none', color: '#dc2626', cursor: 'pointer', padding: '2px' }}
                          title="Eliminar vehículo"
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>
                    ))}
                  </div>

                  {/* Add vehicle form */}
                  <form onSubmit={handleAddVehiculoMant} style={{ display: 'flex', gap: '0.4rem' }}>
                    <input 
                      className="form-input" 
                      placeholder="Placa (ej: HEI-1859)" 
                      value={newVehiculoPlaca} 
                      onChange={(e) => setNewVehiculoPlaca(e.target.value)}
                      style={{ flex: 1, padding: '0.45rem 0.6rem', fontSize: '0.8rem' }}
                    />
                    <select 
                      className="form-input"
                      value={newVehiculoKm}
                      onChange={(e) => setNewVehiculoKm(e.target.value)}
                      style={{ width: '130px', padding: '0.45rem', fontSize: '0.8rem' }}
                    >
                      {[5000, 10000, 15000, 20000, 25000, 30000, 35000, 40000, 45000, 50000, 55000, 60000, 65000, 70000, 75000, 80000, 85000, 90000, 95000, 100000, 105000, 110000, 115000, 120000, 125000, 130000, 140000, 150000, 160000, 170000, 180000, 190000, 200000, 250000].map(k => (
                        <option key={k} value={k}>{k.toLocaleString()} KM</option>
                      ))}
                    </select>
                    <button type="submit" className="btn btn-secondary" style={{ padding: '0.4rem 0.6rem', width: 'auto' }}>
                      <Plus size={14} />
                    </button>
                  </form>
                </div>
              )}

              {/* Range KM (if global range mode active) */}
              {(modoCalculo === 'rango' || (modoCalculo === 'auto' && mantenimientosEspecificos.length === 0)) && (
                <div className="form-group half-width">
                  <div>
                    <label className="form-label">KM Desde</label>
                    <input 
                      type="number"
                      className="form-input" 
                      name="kmDesde" 
                      value={form.kmDesde} 
                      placeholder="5000"
                      onChange={handleChange} 
                    />
                  </div>
                  <div>
                    <label className="form-label">KM Hasta</label>
                    <input 
                      type="number"
                      className="form-input" 
                      name="kmHasta" 
                      value={form.kmHasta} 
                      placeholder="120000"
                      onChange={handleChange} 
                    />
                  </div>
                </div>
              )}

              {/* Plazo & Vigencia */}
              <div className="form-group half-width">
                <div>
                  <label className="form-label">Plazo Ejecución (días)</label>
                  <input 
                    className="form-input" 
                    name="plazoEjecucion" 
                    value={form.plazoEjecucion} 
                    onChange={handleChange} 
                  />
                </div>
                <div>
                  <label className="form-label">Vigencia Oferta</label>
                  <input 
                    className="form-input" 
                    name="vigenciaOferta" 
                    value={form.vigenciaOferta} 
                    onChange={handleChange} 
                  />
                </div>
              </div>

              {/* Placas list */}
              <div className="form-group">
                <label className="form-label">Placas de Vehículos</label>
                <form onSubmit={handleAddPlate} style={{ display: 'flex', gap: '0.5rem' }}>
                  <input 
                    className="form-input" 
                    placeholder="Agregar Placa (Ej: ABA-1234)" 
                    value={newPlate} 
                    onChange={(e) => setNewPlate(e.target.value)} 
                  />
                  <button type="submit" className="btn btn-secondary" style={{ width: 'auto', padding: '0.75rem' }}>
                    <Plus size={16} />
                  </button>
                </form>
                <div className="plates-badge-container">
                  {placas.length > 0 ? (
                    placas.map((plate, index) => (
                      <span className="plate-badge" key={index}>
                        {plate}
                        <button className="plate-badge-remove" onClick={() => handleRemovePlate(index)}>
                          <X size={10} />
                        </button>
                      </span>
                    ))
                  ) : (
                    <span className="upload-subtext">No hay placas ingresadas</span>
                  )}
                </div>
              </div>

              {/* Switch Correctivos */}
              <div className="form-group" style={{ marginBottom: 0 }}>
                <div 
                  className="switch-container" 
                  onClick={() => handleChange({ target: { name: 'incluirCorrectivos', value: !form.incluirCorrectivos, type: 'checkbox', checked: !form.incluirCorrectivos } })}
                >
                  <span className="form-label" style={{ margin: 0 }}>¿Incluir Mantenimiento Correctivo (30%)?</span>
                  <div className={`switch-track ${form.incluirCorrectivos ? 'active' : ''}`}>
                    <div className="switch-thumb"></div>
                  </div>
                </div>
              </div>

            </div>

          </div>

          {/* Right Column: Calculations & Preview */}
          <div className="preview-container">
            
            {/* Summary Widget */}
            <div className="summary-widget">
              <div className="summary-item">
                <span className="summary-label">Mantenimiento Preventivo</span>
                <span className="summary-val">{getFmtCurrency(previewPrev)}</span>
              </div>
              <div className="summary-item" style={{ borderLeft: '1px solid var(--border-light)', paddingLeft: '1.25rem' }}>
                <span className="summary-label">Mantenimiento Correctivo</span>
                <span className="summary-val">{form.incluirCorrectivos ? getFmtCurrency(previewCorr) : '$0.00'}</span>
              </div>
              <div className="summary-item" style={{ borderLeft: '1px solid var(--border-light)', paddingLeft: '1.25rem' }}>
                <span className="summary-label">Gran Total Contrato</span>
                <span className="summary-val total">{getFmtCurrency(previewTotal)}</span>
              </div>
            </div>

            {/* Card: Plan en Uso */}
            <div className="card">
              <h3 className="card-title">
                <FileCode size={20} className="text-muted" />
                Origen de Tarifas & Plan Activo
              </h3>
              
              <div className="alert alert-success" style={{ margin: 0 }}>
                <CheckCircle2 size={18} />
                <div>
                  <strong>{modelosDisponibles.find(m => m.id === selectedModeloId)?.nombre_completo || form.modeloVehiculo || 'Plan Seleccionado'}</strong>
                  <div style={{ fontSize: '0.75rem', marginTop: '2px', opacity: 0.85 }}>
                    {calcResult?.modo === 'individual' ? (
                      <span>Modalidad: Por Vehículo Específico ({calcResult.tablaDetalle.length} Vehículos)</span>
                    ) : (
                      <span>Modalidad: Rango {form.kmDesde || '5000'} KM a {form.kmHasta || '120000'} KM | {form.cantidadVehiculos} Vehículo(s)</span>
                    )}
                  </div>
                </div>
              </div>
            </div>

            {/* Card: Table details preview */}
            <div className="card" style={{ flexGrow: 1, display: 'flex', flexDirection: 'column' }}>
              <h3 className="card-title">
                <FileSpreadsheet size={20} className="text-muted" />
                {calcResult?.modo === 'individual' ? 'Detalle por Vehículo' : 'Detalle de Mantenimiento por KM'}
              </h3>

              {calcError && (
                <div className="alert alert-error" style={{ marginBottom: 0 }}>
                  <AlertTriangle size={18} />
                  <span>{calcError}</span>
                </div>
              )}

              {!calcResult && !calcError && (
                <div className="alert alert-warning" style={{ margin: 0, flexGrow: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  <AlertTriangle size={18} />
                  <span>Carga la cotización o ingresa los datos para ver el desglose.</span>
                </div>
              )}

              {calcResult && !calcError && (
                <div className="table-wrapper" style={{ maxHeight: '400px' }}>
                  <table className="preview-table">
                    <thead>
                      <tr>
                        {calcResult.modo === 'individual' ? (
                          <>
                            <th>Vehículo / Placa</th>
                            <th>Mantenimiento</th>
                            <th className="text-right">Repuestos</th>
                            <th className="text-right">Lubricantes</th>
                            <th className="text-right">Mano de Obra</th>
                            <th className="text-right">Total</th>
                          </>
                        ) : (
                          <>
                            <th>Kilometraje</th>
                            <th className="text-right">Repuestos</th>
                            <th className="text-right">Lubricantes</th>
                            <th className="text-right">Mano de Obra</th>
                            <th className="text-right">Total Unit.</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {calcResult.tablaDetalle.map((row, idx) => (
                        <tr key={idx}>
                          {calcResult.modo === 'individual' ? (
                            <>
                              <td style={{ fontWeight: 600, color: 'var(--primary)' }}>{row.placa}</td>
                              <td>{row.km.toLocaleString()} KM</td>
                              <td className="text-right">{getFmtCurrency(row.repuestos)}</td>
                              <td className="text-right">{getFmtCurrency(row.lubricantes)}</td>
                              <td className="text-right">{getFmtCurrency(row.mano_obra)}</td>
                              <td className="text-right" style={{ fontWeight: 600 }}>{getFmtCurrency(row.total)}</td>
                            </>
                          ) : (
                            <>
                              <td>{row.km.toLocaleString()} KM</td>
                              <td className="text-right">{getFmtCurrency(row.repuestos)}</td>
                              <td className="text-right">{getFmtCurrency(row.lubricantes)}</td>
                              <td className="text-right">{getFmtCurrency(row.mano_obra)}</td>
                              <td className="text-right" style={{ fontWeight: 600 }}>{getFmtCurrency(row.total)}</td>
                            </>
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* Action Card: Generate docx */}
            <div className="card" style={{ padding: '1.25rem' }}>
              <button 
                className="btn btn-primary"
                disabled={!calcResult || isGenerating}
                onClick={handleGenerateWord}
              >
                <Download size={18} />
                {isGenerating ? 'Generando y Guardando en Supabase...' : 'Descargar Proforma (.docx)'}
              </button>
            </div>

          </div>

        </div>
      ) : (
        /* History View */
        <div className="card" style={{ minHeight: '450px' }}>
          <h3 className="card-title">
            <History size={20} className="text-muted" />
            Historial de Proformas en Supabase
          </h3>

          {loadingHistory ? (
            <div style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-muted)' }}>
              Cargando historial de proformas...
            </div>
          ) : historialProformas.length === 0 ? (
            <div className="alert alert-warning">
              <AlertTriangle size={18} />
              <span>Aún no hay proformas guardadas en la base de datos de Supabase.</span>
            </div>
          ) : (
            <div className="table-wrapper">
              <table className="preview-table">
                <thead>
                  <tr>
                    <th>Nro Proforma</th>
                    <th>Cliente / Razón Social</th>
                    <th>Modelo</th>
                    <th className="text-center">Cant.</th>
                    <th>Rango KM</th>
                    <th className="text-right">Gran Total</th>
                    <th>Fecha</th>
                  </tr>
                </thead>
                <tbody>
                  {historialProformas.map((item, idx) => (
                    <tr key={item.id || idx}>
                      <td style={{ fontWeight: 600, color: 'var(--primary)' }}>{item.numero_proforma}</td>
                      <td>{item.razon_social || 'N/A'}</td>
                      <td>{item.modelo_vehiculo || 'N/A'}</td>
                      <td className="text-center">{item.cantidad_vehiculos}</td>
                      <td>{item.km_desde?.toLocaleString()} - {item.km_hasta?.toLocaleString()} KM</td>
                      <td className="text-right" style={{ fontWeight: 700 }}>{getFmtCurrency(item.gran_total)}</td>
                      <td>{new Date(item.created_at).toLocaleDateString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
      
      {/* Footer credits */}
      <footer className="footer-credits">
        AMBACAR PROFORMAS &copy; {new Date().getFullYear()}. Base de datos en Supabase PostgreSQL.
      </footer>
    </div>
  );
}
