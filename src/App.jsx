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
  X
} from 'lucide-react';
import { 
  buscarValor, 
  parseRangoKM, 
  detectarCorrectivosExcel, 
  parsePlacas, 
  calcularPlanMantenimiento 
} from './utils/excelParser';
import { generarProformaDocx } from './utils/docxGenerator';
import { 
  getModelosVehiculo, 
  buscarPlanPorModelo, 
  getCostosPorKm, 
  guardarProforma, 
  getHistorialProformas 
} from './services/supabase';

export default function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState('generator'); // 'generator' | 'history'

  // Files State
  const [cotizacionFile, setCotizacionFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [showManualTemplate, setShowManualTemplate] = useState(false);
  
  // Supabase State
  const [modelosDisponibles, setModelosDisponibles] = useState([]);
  const [selectedPlanId, setSelectedPlanId] = useState("");
  const [selectedModeloId, setSelectedModeloId] = useState("");
  const [supabaseConnected, setSupabaseConnected] = useState(false);
  const [historialProformas, setHistorialProformas] = useState([]);
  const [loadingHistory, setLoadingHistory] = useState(false);

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
    modeloVehiculo: "",
    cantidadVehiculos: 1,
    plazoEjecucion: "365",
    vigenciaOferta: "90 días",
    observacion: "",
    objetoContrato: "Plan de mantenimiento preventivo y correctivo de vehículos",
    kmDesde: "",
    kmHasta: "",
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
      const { data, error } = await getModelosVehiculo();
      if (!error && data && data.length > 0) {
        setModelosDisponibles(data);
        setSupabaseConnected(true);
        // Default to first model plan
        if (data[0].planes_mantenimiento && data[0].planes_mantenimiento.length > 0) {
          setSelectedModeloId(data[0].id);
          setSelectedPlanId(data[0].planes_mantenimiento[0].id);
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

  // Parse Quotation File (DATOS_PARA_COTIZACION.xlsx)
  const handleCotizacionChange = (e) => {
    const file = e.target.files[0];
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
        
        // Extract fields
        const razonSocialVal = buscarValor(grid, "Razón Social") || "";
        const rucVal = buscarValor(grid, "RUC") || "";
        const direccionVal = buscarValor(grid, "Dirección") || "";
        const contactoVal = buscarValor(grid, "Nombre del Contacto") || "";
        const telefonoVal = buscarValor(grid, "Teléfono") || "";
        const correoVal = buscarValor(grid, "Correo electrónico") || "";
        const plazoVal = buscarValor(grid, "Plazo de ejecución") || "365";
        const vigenciaVal = buscarValor(grid, "Vigencia de la oferta") || "90 días";
        const modeloVal = buscarValor(grid, "Modelo de vehículo") || "";
        const cantidadVal = buscarValor(grid, "Cantidad de vehículos") || "1";
        const observacionVal = buscarValor(grid, "Observación") || "";
        const objetoVal = buscarValor(grid, "OBJETO DEL CONTRATO") || "Plan de mantenimiento preventivo y correctivo de vehículos";
        
        const { kmDesde: kd, kmHasta: kh } = parseRangoKM(grid);
        const placasVal = parsePlacas(grid);
        
        // Detect checkboxes using zip vml
        let inclCorr = false;
        const zipResult = await detectarCorrectivosExcel(data, grid);
        if (zipResult) {
          inclCorr = zipResult.siChecked && !zipResult.noChecked;
        } else {
          // Fallback text check
          for (let r = 0; r < grid.length; r++) {
            const row = grid[r] || [];
            const c1 = row[1]?.toString().trim() || "";
            if (c1.toLowerCase().includes("correctivo") && c1.toLowerCase().includes("incluir")) {
              for (const vc of [3, 4, 5]) {
                if (vc < row.length) {
                  const v = row[vc]?.toString().trim().toUpperCase();
                  if (v === "SI" || v === "SÍ") inclCorr = true;
                }
              }
            }
          }
        }

        setForm(prev => ({
          ...prev,
          razonSocial: razonSocialVal,
          ruc: rucVal,
          direccion: direccionVal,
          contacto: contactoVal,
          telefono: telefonoVal,
          correo: correoVal,
          plazoEjecucion: plazoVal,
          vigenciaOferta: vigenciaVal,
          modeloVehiculo: modeloVal,
          cantidadVehiculos: cantidadVal,
          observacion: observacionVal,
          objetoContrato: objetoVal,
          kmDesde: kd || prev.kmDesde || "5000",
          kmHasta: kh || prev.kmHasta || "120000",
          incluirCorrectivos: inclCorr
        }));
        
        setPlacas(placasVal);

        // Match with Supabase Model & Plan
        if (modeloVal) {
          const { plan, modelo } = await buscarPlanPorModelo(modeloVal);
          if (plan && modelo) {
            setSelectedModeloId(modelo.id);
            setSelectedPlanId(plan.id);
          }
        }
      } catch (err) {
        console.error("Error reading quotation file:", err);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Parse Maintenance Plan Template File (Offline Manual Fallback)
  const handleTemplateChange = (e) => {
    const file = e.target.files[0];
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

  // Recalculate cost when range, vehicles, plan, or manual sheet changes
  useEffect(() => {
    const kDesde = parseInt(form.kmDesde, 10);
    const kHasta = parseInt(form.kmHasta, 10);
    const nVehic = parseInt(form.cantidadVehiculos, 10) || 1;

    if (isNaN(kDesde) || isNaN(kHasta) || kDesde <= 0 || kHasta <= 0) {
      return;
    }

    // 1. If using Supabase Plan (Default Mode)
    if (selectedPlanId && !templateWorkbook) {
      getCostosPorKm(selectedPlanId, kDesde, kHasta, nVehic).then(res => {
        if (res.error) {
          setCalcResult(null);
          setCalcError(res.message || "Error al calcular desde Supabase");
        } else {
          setCalcResult(res);
          setCalcError(null);
        }
      });
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
  }, [selectedPlanId, templateWorkbook, sheetSelection, form.kmDesde, form.kmHasta, form.cantidadVehiculos]);

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
    
    const totalRepuestos = calcResult.totalRepuestos;
    const totalLubricantes = calcResult.totalLubricantes;
    const totalManoObra = calcResult.totalMo;
    const totalPreventivo = calcResult.totalPreventivo;
    const totalCorrectivo = form.incluirCorrectivos ? Math.round(totalPreventivo * 0.3 * 100) / 100 : 0;
    const granTotal = totalPreventivo + totalCorrectivo;

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

  const previewPrev = calcResult ? calcResult.totalPreventivo : 0;
  const previewCorr = form.incluirCorrectivos ? Math.round(previewPrev * 0.3 * 100) / 100 : 0;
  const previewTotal = previewPrev + previewCorr;

  return (
    <div className="app-container">
      {/* Header section */}
      <header className="app-header">
        <div className="brand-section">
          <span className="brand-logo">🚗</span>
          <div>
            <h1 className="brand-title">Generador de Proformas AMBACAR</h1>
            <p className="brand-subtitle">Cotizaciones conectadas con Base de Datos Supabase</p>
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
                  <div className="upload-zone" onClick={() => document.getElementById('cot-input').click()}>
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
                      <div className="upload-zone" style={{ padding: '1rem' }} onClick={() => document.getElementById('template-input').click()}>
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
                    <div className="upload-zone" style={{ padding: '1rem' }} onClick={() => document.getElementById('header-img-input').click()}>
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
                    <div className="upload-zone" style={{ padding: '1rem' }} onClick={() => document.getElementById('footer-img-input').click()}>
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
                    placeholder="Ej: Wingle 7 4x4"
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

              {/* Range KM */}
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
                    Rango: {form.kmDesde || '5000'} KM a {form.kmHasta || '120000'} KM | {form.cantidadVehiculos} Vehículo(s)
                  </div>
                </div>
              </div>
            </div>

            {/* Card: Table details preview */}
            <div className="card" style={{ flexGrow: 1, display: 'flex', flexDirection: 'column' }}>
              <h3 className="card-title">
                <FileSpreadsheet size={20} className="text-muted" />
                Detalle de Mantenimiento por KM
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
                  <span>Carga la cotización o ingresa los kilómetros para ver el desglose.</span>
                </div>
              )}

              {calcResult && !calcError && (
                <div className="table-wrapper" style={{ maxHeight: '400px' }}>
                  <table className="preview-table">
                    <thead>
                      <tr>
                        <th>Kilometraje</th>
                        <th className="text-right">Repuestos</th>
                        <th className="text-right">Lubricantes</th>
                        <th className="text-right">Mano de Obra</th>
                        <th className="text-right">Total Unit.</th>
                      </tr>
                    </thead>
                    <tbody>
                      {calcResult.tablaDetalle.map((row, idx) => (
                        <tr key={idx}>
                          <td>{row.km.toLocaleString()} KM</td>
                          <td className="text-right">{getFmtCurrency(row.repuestos)}</td>
                          <td className="text-right">{getFmtCurrency(row.lubricantes)}</td>
                          <td className="text-right">{getFmtCurrency(row.mano_obra)}</td>
                          <td className="text-right" style={{ fontWeight: 600 }}>{getFmtCurrency(row.total)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* Action Card: Generate docx */}
            <div className="card" style={{ padding: '1rem', border: '1px solid rgba(255,255,255,0.1)' }}>
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
        Generador de Proformas Ambacar &copy; {new Date().getFullYear()}. Base de datos en Supabase PostgreSQL.
      </footer>
    </div>
  );
}
