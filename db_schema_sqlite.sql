-- PV Elite Extraction Database Schema (SQLite)
--
-- Normalized design: one row per TAG in equipment_master, with proper
-- one-to-many child tables for elements/platforms/weight-breakdown,
-- rather than the wide, duplicate-column-across-sheets layout used in
-- the Excel output (that layout is fine for a human reading Excel;
-- it's the wrong shape for a real database).
--
-- No server, no install - this file itself IS the database. For
-- multiple users to accumulate into the same database over time, point
-- everyone's app at the same file on a shared network drive.

PRAGMA foreign_keys = ON;

-- One row per TAG. project + tag_no together uniquely identify a
-- vessel, so re-importing the same tag UPDATES this row (see the
-- upsert logic in db_writer.py) instead of creating a duplicate.
CREATE TABLE IF NOT EXISTS equipment_master (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    project             TEXT NOT NULL,
    category            TEXT NOT NULL,
    tag_no              TEXT NOT NULL,
    equipment_name      TEXT,
    file_kind           TEXT,
    warnings_count      INTEGER,

    orientation         TEXT,
    support_type        TEXT,
    asme_code           TEXT,
    division            INTEGER,
    code_edition_year   TEXT,
    mdmt_c              REAL,

    shell_moc           TEXT,
    shell_moc_class     TEXT,
    shell_uns           TEXT,

    tl_tl_mm                 REAL,
    max_diameter_mm           REAL,
    min_diameter_mm            REAL,
    l_over_d_ratio               REAL,

    max_design_pressure_mpa       REAL,
    max_design_temp_c               REAL,

    wind_code                          TEXT,
    seismic_code                         TEXT,
    basic_wind_speed_kmh                   REAL,
    site_class                               TEXT,
    seismic_ss                                 REAL,
    seismic_s1                                   REAL,
    seismic_sds                                    REAL,
    seismic_sd1                                      REAL,
    seismic_response_coeff_cs                          REAL,
    max_wind_pressure_kgm2                               REAL,

    wind_shear_kgf       REAL,
    eq_shear_kgf           REAL,
    wind_moment_kgfm         REAL,
    eq_moment_kgfm             REAL,

    fabricated_mt        REAL,
    shop_test_mt           REAL,
    shipping_mt              REAL,
    erected_mt                 REAL,
    empty_mt                    REAL,
    operating_mt                  REAL,
    total_element_weight_mt         REAL,
    total_surface_area_m2             REAL,

    has_platform          INTEGER,  -- 0/1 (SQLite has no native boolean)
    has_insulation           INTEGER,
    has_fp                     INTEGER,
    total_platform_area_m2       REAL,
    insulation_area_m2             REAL,
    painting_area_m2                  REAL,
    fp_area_m2                          REAL,

    basering_type         TEXT,
    bolt_moc                TEXT,
    bolt_nominal_dia_mm       REAL,
    bolt_circle_dia_mm          REAL,
    bolt_qty                      INTEGER,

    imported_by            TEXT,
    imported_at              TEXT DEFAULT CURRENT_TIMESTAMP,
    source_file              TEXT,

    UNIQUE (project, tag_no)
);

-- One row per ELEMENT (many per tag). Deleted and re-inserted on
-- re-import of the same tag so element counts never drift if a tag's
-- report changes between imports.
CREATE TABLE IF NOT EXISTS element_data (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id        INTEGER NOT NULL,
    element_no           INTEGER,
    element_name           TEXT,
    component                TEXT,
    material                   TEXT,
    class                        TEXT,
    uns_number                     TEXT,
    diameter_m                       REAL,
    element_length_mm                  REAL,
    minimum_thickness_mm                 REAL,
    nominal_thickness_mm                   REAL,
    corrosion_allowance_mm                   REAL,
    external_pressure_mpa                      REAL,
    design_pressure_mpa                          REAL,
    design_temperature_c                           REAL,
    allowable_tensile_stress_mpa                     REAL,
    allowable_compressive_stress_mpa                   REAL,
    surface_area_m2                                      REAL,
    inside_surface_area_m2                                 REAL,
    weight_mt                                                REAL,
    total_ele_empty_wgt_mt                                     REAL,

    FOREIGN KEY (equipment_id) REFERENCES equipment_master(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_element_equipment ON element_data(equipment_id);

-- One row per PLATFORM (many per tag).
CREATE TABLE IF NOT EXISTS platform_data (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id        INTEGER NOT NULL,
    element_name          TEXT,
    diameter_mm             REAL,
    detail_id                 TEXT,
    start_angle_deg             REAL,
    end_angle_deg                 REAL,
    width_mm                        REAL,
    length_mm                         REAL,
    area_m2                             REAL,

    FOREIGN KEY (equipment_id) REFERENCES equipment_master(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_platform_equipment ON platform_data(equipment_id);

-- One row per (tag, weight-summation category) - matches the
-- "weight_summ" Excel sheet. "..." (not applicable) is stored as NULL,
-- the correct database representation - "..." was purely an Excel-
-- display convenience.
CREATE TABLE IF NOT EXISTS weight_summation (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id        INTEGER NOT NULL,
    component              TEXT,
    fabricated_kg             REAL,
    shop_test_kg                REAL,
    shipping_kg                   REAL,
    erected_kg                      REAL,
    empty_kg                          REAL,
    operating_kg                        REAL,

    FOREIGN KEY (equipment_id) REFERENCES equipment_master(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_weightsumm_equipment ON weight_summation(equipment_id);

-- ============================================================================
-- SCHEMA v2 additions (additive only - existing v1 tables/columns above are
-- untouched, so any database already running v1 upgrades safely in place).
-- Adds: equipment-type generality (so this schema covers Heat Exchangers
-- later without restructuring), a flexible EAV table for mechanical
-- properties that vary by equipment type, and a documents table holding
-- each report's per-section text - the source of truth the vector DB's
-- embeddings are built from, and what lets you re-embed later if the
-- embedding model ever changes without re-parsing the original PDFs.
-- ============================================================================

CREATE TABLE IF NOT EXISTS schema_version (
    version     INTEGER PRIMARY KEY,
    applied_at  TEXT DEFAULT CURRENT_TIMESTAMP,
    notes       TEXT
);
INSERT OR IGNORE INTO schema_version (version, notes) VALUES
    (1, 'Initial frozen schema: equipment_master, element_data, platform_data, weight_summation'),
    (2, 'Added equipment_type generality, mechanical_details (EAV), documents (vector-DB source text)');

-- equipment_type distinguishes Vessel/Column (built now) from
-- HeatExchanger (schema-ready, extraction to be built later) without
-- needing a different table per type. Existing rows default to
-- 'Vessel' since that's everything imported under schema v1.
ALTER TABLE equipment_master ADD COLUMN equipment_type TEXT DEFAULT 'Vessel';

-- Generic mechanical property table (Component/Property/Value/Units),
-- deliberately NOT a fixed column per property - a vessel's relevant
-- properties (Skirt Thickness, Basering...) and a heat exchanger's
-- (Tubesheet Thickness, Channel Flange...) barely overlap, so a wide
-- fixed-column table would mean either constant schema changes or a
-- sea of always-NULL columns. This table is populated for
-- Vessels/Columns now (rolled up from element_data - max nominal
-- thickness per Component, matching the "display maximum value for
-- each category" mechanical-summary convention) and will hold the
-- Heat Exchanger equivalents once that extraction is built, without
-- any schema change needed at that point.
CREATE TABLE IF NOT EXISTS mechanical_details (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id INTEGER NOT NULL,
    component    TEXT NOT NULL,
    property     TEXT NOT NULL,
    value        REAL,
    units        TEXT,

    FOREIGN KEY (equipment_id) REFERENCES equipment_master(id) ON DELETE CASCADE,
    UNIQUE (equipment_id, component, property)
);
CREATE INDEX IF NOT EXISTS idx_mechdetails_equipment ON mechanical_details(equipment_id);

-- One row per (tag, report section) - "Vessel Design Summary", "Nozzle
-- Summary", "MDMT Summary", etc. This is the vector DB's source
-- material: each row here becomes (or re-becomes) one chunk embedded
-- into the vector store. Keeping the raw text in SQL means changing
-- the embedding model later is a re-embed of what's already here, not
-- a re-parse of the original PDFs.
CREATE TABLE IF NOT EXISTS documents (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    equipment_id   INTEGER NOT NULL,
    section_title  TEXT NOT NULL,
    section_text   TEXT,
    embedded       INTEGER DEFAULT 0,  -- 0/1: has this row been pushed to the vector DB yet
    embedded_at    TEXT,

    FOREIGN KEY (equipment_id) REFERENCES equipment_master(id) ON DELETE CASCADE,
    UNIQUE (equipment_id, section_title)
);
CREATE INDEX IF NOT EXISTS idx_documents_equipment ON documents(equipment_id);
CREATE INDEX IF NOT EXISTS idx_documents_unembedded ON documents(embedded) WHERE embedded = 0;
