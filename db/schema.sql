-- Mesa de Entradas - esquema PostgreSQL (Supabase)
-- Todo vive en el esquema mesa_entradas, separado de las tablas de viáticos.
-- Para que la app lo vea por la API: Supabase > Project Settings > API > Exposed schemas,
-- agregar mesa_entradas.

CREATE SCHEMA IF NOT EXISTS mesa_entradas;
SET search_path TO mesa_entradas;

-- ---------------------------------------------------------------
-- Catálogos
-- ---------------------------------------------------------------

CREATE TABLE IF NOT EXISTS areas (
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nombre TEXT NOT NULL UNIQUE,
    activa BOOLEAN NOT NULL DEFAULT true
);

CREATE TABLE IF NOT EXISTS organismos (
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nombre TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS responsables (
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nombre TEXT NOT NULL,
    area_id BIGINT REFERENCES areas(id)
);

-- ---------------------------------------------------------------
-- Numeración propia de la mesa: un contador por tipo y año
-- ---------------------------------------------------------------

CREATE TABLE IF NOT EXISTS numeradores (
    tipo TEXT NOT NULL CHECK (tipo IN ('Entrada', 'Salida')),
    anio SMALLINT NOT NULL,
    ultimo INT NOT NULL DEFAULT 0,
    PRIMARY KEY (tipo, anio)
);

-- Si no viene número, asigna el siguiente. Si viene a mano (carga de
-- registros viejos), lo respeta y adelanta el contador para no repetirlo.
-- Las entradas por correo no llevan número de mesa.
CREATE OR REPLACE FUNCTION asignar_numero() RETURNS trigger AS $$
BEGIN
    IF TG_TABLE_NAME = 'entradas' THEN
        IF NEW.medio = 'Correo' THEN
            RETURN NEW;
        END IF;
    END IF;
    IF NEW.numero IS NULL THEN
        INSERT INTO numeradores (tipo, anio, ultimo)
        VALUES (TG_ARGV[0], NEW.anio, 1)
        ON CONFLICT (tipo, anio) DO UPDATE SET ultimo = numeradores.ultimo + 1
        RETURNING ultimo INTO NEW.numero;
    ELSE
        INSERT INTO numeradores (tipo, anio, ultimo)
        VALUES (TG_ARGV[0], NEW.anio, NEW.numero)
        ON CONFLICT (tipo, anio) DO UPDATE
            SET ultimo = GREATEST(numeradores.ultimo, EXCLUDED.ultimo);
    END IF;
    RETURN NEW;
END $$ LANGUAGE plpgsql SET search_path = mesa_entradas;

CREATE OR REPLACE FUNCTION set_updated_at() RETURNS trigger AS $$
BEGIN
    NEW.updated_at = now();
    RETURN NEW;
END $$ LANGUAGE plpgsql SET search_path = mesa_entradas;

-- ---------------------------------------------------------------
-- Entradas
-- ---------------------------------------------------------------

CREATE TABLE IF NOT EXISTS entradas (
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    anio SMALLINT NOT NULL DEFAULT EXTRACT(YEAR FROM CURRENT_DATE),
    medio TEXT NOT NULL DEFAULT 'Papel' CHECK (medio IN ('Papel', 'Correo')),
    numero INT,                                       -- de mesa: solo Papel, lo asigna el trigger
    clase TEXT CHECK (clase IN ('Oficio', 'Nota', 'Expediente', 'Memo', 'Otro')),
    numero_documento TEXT,                            -- número de la nota/oficio recibido
    numero_expediente TEXT,
    fecha_documento DATE,
    fecha_recepcion DATE NOT NULL DEFAULT CURRENT_DATE,
    organismo_id BIGINT REFERENCES organismos(id),    -- quién lo envía (catálogo)
    remitente TEXT,                                   -- quién lo envía, tal como figura
    area_id BIGINT REFERENCES areas(id),              -- a qué área va
    firmantes TEXT,
    asunto TEXT NOT NULL,
    extracto TEXT,
    fecha_vencimiento DATE,                           -- según el oficio; vacío si no tiene plazo
    responsable_id BIGINT REFERENCES responsables(id),
    estado TEXT NOT NULL DEFAULT 'Pendiente'
        CHECK (estado IN ('Pendiente', 'En Trámite', 'Respondido', 'Archivado')),
    observaciones TEXT,
    ruta_archivo TEXT UNIQUE,
    correo_id TEXT UNIQUE,                            -- EntryID de Outlook (medio Correo)
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (anio, numero)
);

CREATE INDEX IF NOT EXISTS idx_entradas_num_doc ON entradas (numero_documento);
CREATE INDEX IF NOT EXISTS idx_entradas_expte ON entradas (numero_expediente);
CREATE INDEX IF NOT EXISTS idx_entradas_venc ON entradas (estado, fecha_vencimiento);

CREATE OR REPLACE TRIGGER trg_entradas_numero BEFORE INSERT ON entradas
FOR EACH ROW EXECUTE FUNCTION asignar_numero('Entrada');

CREATE OR REPLACE TRIGGER trg_entradas_updated BEFORE UPDATE ON entradas
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

-- ---------------------------------------------------------------
-- Salidas
-- ---------------------------------------------------------------

CREATE TABLE IF NOT EXISTS salidas (
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    anio SMALLINT NOT NULL DEFAULT EXTRACT(YEAR FROM CURRENT_DATE),
    numero INT NOT NULL,                              -- lo asigna el trigger
    sufijo TEXT NOT NULL DEFAULT '',                  -- b, c...: notas viejas que repitieron número
    clase TEXT CHECK (clase IN ('Oficio', 'Nota', 'Expediente', 'Memo', 'Otro')),
    numero_documento TEXT,                            -- como figura en la nota, ej. 0137/25
    numero_expediente TEXT,
    fecha_salida DATE DEFAULT CURRENT_DATE,           -- vacía si una nota vieja no la trae
    area_id BIGINT REFERENCES areas(id),              -- área que la emite
    organismo_id BIGINT REFERENCES organismos(id),    -- a quién va (catálogo)
    destinatario TEXT,                                -- a quién va, tal como figura
    firmantes TEXT,
    asunto TEXT NOT NULL,
    extracto TEXT,
    entrada_id BIGINT REFERENCES entradas(id),        -- entrada que responde, si corresponde
    observaciones TEXT,
    ruta_archivo TEXT UNIQUE,
    correo_id TEXT UNIQUE,                            -- EntryID de Outlook, si salió por correo
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (anio, numero, sufijo)
);

CREATE INDEX IF NOT EXISTS idx_salidas_num_doc ON salidas (numero_documento);
CREATE INDEX IF NOT EXISTS idx_salidas_expte ON salidas (numero_expediente);
CREATE INDEX IF NOT EXISTS idx_salidas_entrada ON salidas (entrada_id);

CREATE OR REPLACE TRIGGER trg_salidas_numero BEFORE INSERT ON salidas
FOR EACH ROW EXECUTE FUNCTION asignar_numero('Salida');

CREATE OR REPLACE TRIGGER trg_salidas_updated BEFORE UPDATE ON salidas
FOR EACH ROW EXECUTE FUNCTION set_updated_at();

-- Al registrar una salida que responde a una entrada, la entrada pasa a Respondido.
CREATE OR REPLACE FUNCTION marcar_respondida() RETURNS trigger AS $$
BEGIN
    IF NEW.entrada_id IS NOT NULL THEN
        UPDATE entradas SET estado = 'Respondido'
        WHERE id = NEW.entrada_id AND estado IN ('Pendiente', 'En Trámite');
    END IF;
    RETURN NEW;
END $$ LANGUAGE plpgsql SET search_path = mesa_entradas;

CREATE OR REPLACE TRIGGER trg_salidas_respuesta AFTER INSERT OR UPDATE OF entrada_id ON salidas
FOR EACH ROW EXECUTE FUNCTION marcar_respondida();

-- ---------------------------------------------------------------
-- Vista de entradas con respuesta y vencimiento
-- ---------------------------------------------------------------

CREATE OR REPLACE VIEW v_entradas AS
SELECT e.*,
       r.salidas IS NOT NULL AS respondida,
       r.salidas AS respuestas,          -- ej: '0015/26, 0043/26'
       COALESCE(e.estado IN ('Pendiente', 'En Trámite')
                AND e.fecha_vencimiento < CURRENT_DATE, false) AS vencida
FROM entradas e
LEFT JOIN (
    SELECT entrada_id,
           string_agg(COALESCE(numero_documento, numero || '/' || anio)
                      || CASE WHEN sufijo <> '' THEN ' ' || sufijo ELSE '' END,
                      ', ' ORDER BY anio, numero, sufijo) AS salidas
    FROM salidas
    WHERE entrada_id IS NOT NULL
    GROUP BY entrada_id
) r ON r.entrada_id = e.id;

-- ---------------------------------------------------------------
-- Permisos: solo la app (service_role). anon y authenticated no ven nada.
-- ---------------------------------------------------------------

GRANT USAGE ON SCHEMA mesa_entradas TO service_role;
GRANT ALL ON ALL TABLES IN SCHEMA mesa_entradas TO service_role;
GRANT ALL ON ALL SEQUENCES IN SCHEMA mesa_entradas TO service_role;
GRANT EXECUTE ON ALL FUNCTIONS IN SCHEMA mesa_entradas TO service_role;
ALTER DEFAULT PRIVILEGES IN SCHEMA mesa_entradas GRANT ALL ON TABLES TO service_role;
ALTER DEFAULT PRIVILEGES IN SCHEMA mesa_entradas GRANT ALL ON SEQUENCES TO service_role;
