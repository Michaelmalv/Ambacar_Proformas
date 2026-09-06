-- ==============================================================================
-- ESQUEMA DE BASE DE DATOS PARA GENERADOR DE PROFORMAS AMBACAR (SUPABASE / POSTGRESQL)
-- ==============================================================================

-- Habilitar extensión para UUIDs
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 1. TABLA: MODELOS DE VEHÍCULOS
CREATE TABLE IF NOT EXISTS public.modelos_vehiculo (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    marca VARCHAR(50) NOT NULL,            -- ej. 'GWM', 'HAVAL', 'TANK', 'KYC'
    modelo VARCHAR(100) NOT NULL,          -- ej. 'WINGLE 7', 'POER', 'TANK 300', 'TANK 500', 'KYC F3', 'WINGLE 2.8'
    motor VARCHAR(50) DEFAULT 'DIESEL',    -- ej. 'DIESEL', 'GASOLINA', '2.0 DIESEL'
    traccion VARCHAR(50) DEFAULT '4X4',    -- ej. '4X4', '4X2', '4X4 Y 4X2'
    nombre_completo VARCHAR(150) NOT NULL UNIQUE, -- ej. 'WINGLE 7 DIESEL 4X4'
    categoria VARCHAR(50) DEFAULT 'CAMIONETA',
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. TABLA: PLANES DE MANTENIMIENTO PREVENTIVO
CREATE TABLE IF NOT EXISTS public.planes_mantenimiento (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    vehiculo_id UUID REFERENCES public.modelos_vehiculo(id) ON DELETE CASCADE,
    codigo_plan VARCHAR(100) NOT NULL UNIQUE,
    nombre VARCHAR(200) NOT NULL,
    archivo_origen VARCHAR(255),
    hoja_origen VARCHAR(100),
    km_inicio INT DEFAULT 5000,
    km_fin INT DEFAULT 120000,
    intervalo_km INT DEFAULT 5000,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. TABLA: COSTOS TOTALES POR KILOMETRAJE (1 VEHÍCULO)
CREATE TABLE IF NOT EXISTS public.plan_costos_km (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    plan_id UUID REFERENCES public.planes_mantenimiento(id) ON DELETE CASCADE,
    km INT NOT NULL,
    total_repuestos NUMERIC(12, 2) NOT NULL DEFAULT 0.00,
    total_lubricantes NUMERIC(12, 2) NOT NULL DEFAULT 0.00,
    total_mano_obra NUMERIC(12, 2) NOT NULL DEFAULT 0.00,
    total_km NUMERIC(12, 2) NOT NULL DEFAULT 0.00,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    CONSTRAINT uq_plan_km UNIQUE(plan_id, km)
);

-- 4. TABLA: DESGLOSE DE ÍTEMS Y RUBROS POR KILOMETRAJE
CREATE TABLE IF NOT EXISTS public.plan_items_detalle (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    plan_id UUID REFERENCES public.planes_mantenimiento(id) ON DELETE CASCADE,
    categoria VARCHAR(50) NOT NULL, -- 'LUBRICANTES', 'REPUESTOS', 'MANO DE OBRA'
    descripcion VARCHAR(255) NOT NULL,
    km INT NOT NULL,
    costo NUMERIC(12, 2) DEFAULT 0.00,
    aplica BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. TABLA: CATÁLOGO DE MANTENIMIENTO CORRECTIVO
CREATE TABLE IF NOT EXISTS public.catalogo_correctivo (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    vehiculo_id UUID REFERENCES public.modelos_vehiculo(id) ON DELETE SET NULL,
    modelo_referencia VARCHAR(100) NOT NULL, -- ej. 'WINGLE 7 DIESEL', 'POER DIESEL', 'TANK 300'
    tipo_catalogo VARCHAR(50) DEFAULT 'ESTANDAR', -- 'ESTANDAR', 'COMPLETO', 'SOLICITADOS'
    item_numero INT,
    descripcion VARCHAR(255) NOT NULL,
    costo_repuestos NUMERIC(12, 2) DEFAULT 0.00,
    costo_mano_obra NUMERIC(12, 2) DEFAULT 0.00,
    costo_lubricantes NUMERIC(12, 2) DEFAULT 0.00,
    subtotal NUMERIC(12, 2) DEFAULT 0.00,
    iva NUMERIC(12, 2) DEFAULT 0.00,
    total NUMERIC(12, 2) DEFAULT 0.00,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 6. TABLA: HISTORIAL DE PROFORMAS GENERADAS
CREATE TABLE IF NOT EXISTS public.proformas_generadas (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    numero_proforma VARCHAR(100) NOT NULL UNIQUE,
    razon_social VARCHAR(255),
    ruc VARCHAR(20),
    direccion TEXT,
    contacto VARCHAR(255),
    telefono VARCHAR(100),
    correo VARCHAR(255),
    objeto_contrato TEXT,
    modelo_vehiculo VARCHAR(150),
    cantidad_vehiculos INT DEFAULT 1,
    km_desde INT,
    km_hasta INT,
    incluye_correctivos BOOLEAN DEFAULT FALSE,
    total_repuestos NUMERIC(14, 2) DEFAULT 0.00,
    total_lubricantes NUMERIC(14, 2) DEFAULT 0.00,
    total_mano_obra NUMERIC(14, 2) DEFAULT 0.00,
    total_preventivo NUMERIC(14, 2) DEFAULT 0.00,
    total_correctivo NUMERIC(14, 2) DEFAULT 0.00,
    gran_total NUMERIC(14, 2) DEFAULT 0.00,
    placas TEXT[],
    detalle_km JSONB,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- ==============================================================================
-- ÍNDICES PARA BÚSQUEDAS RÁPIDAS
-- ==============================================================================
CREATE INDEX IF NOT EXISTS idx_planes_vehiculo ON public.planes_mantenimiento(vehiculo_id);
CREATE INDEX IF NOT EXISTS idx_plan_costos_km ON public.plan_costos_km(plan_id, km);
CREATE INDEX IF NOT EXISTS idx_plan_items ON public.plan_items_detalle(plan_id, km);
CREATE INDEX IF NOT EXISTS idx_correctivos_modelo ON public.catalogo_correctivo(modelo_referencia);
CREATE INDEX IF NOT EXISTS idx_proformas_numero ON public.proformas_generadas(numero_proforma);

-- ==============================================================================
-- POLÍTICAS DE SEGURIDAD (ROW LEVEL SECURITY - RLS)
-- Permite lectura pública de catálogos y precios, e inserción de proformas
-- ==============================================================================
ALTER TABLE public.modelos_vehiculo ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.planes_mantenimiento ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.plan_costos_km ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.plan_items_detalle ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.catalogo_correctivo ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.proformas_generadas ENABLE ROW LEVEL SECURITY;

-- Políticas de lectura (SELECT) para todos (anon y authenticated)
CREATE POLICY "Permitir lectura publica de modelos_vehiculo" ON public.modelos_vehiculo FOR SELECT USING (true);
CREATE POLICY "Permitir lectura publica de planes_mantenimiento" ON public.planes_mantenimiento FOR SELECT USING (true);
CREATE POLICY "Permitir lectura publica de plan_costos_km" ON public.plan_costos_km FOR SELECT USING (true);
CREATE POLICY "Permitir lectura publica de plan_items_detalle" ON public.plan_items_detalle FOR SELECT USING (true);
CREATE POLICY "Permitir lectura publica de catalogo_correctivo" ON public.catalogo_correctivo FOR SELECT USING (true);
CREATE POLICY "Permitir lectura publica de proformas_generadas" ON public.proformas_generadas FOR SELECT USING (true);

-- Políticas de inserción y modificación
CREATE POLICY "Permitir insercion publica de proformas_generadas" ON public.proformas_generadas FOR INSERT WITH CHECK (true);
CREATE POLICY "Permitir insercion de modelos para service role" ON public.modelos_vehiculo FOR ALL USING (true);
CREATE POLICY "Permitir insercion de planes para service role" ON public.planes_mantenimiento FOR ALL USING (true);
CREATE POLICY "Permitir insercion de costos para service role" ON public.plan_costos_km FOR ALL USING (true);
CREATE POLICY "Permitir insercion de items para service role" ON public.plan_items_detalle FOR ALL USING (true);
CREATE POLICY "Permitir insercion de correctivos para service role" ON public.catalogo_correctivo FOR ALL USING (true);
