"""
db_writer.py

Writes PVEliteDynamicExtractor results into the SQLite database defined
in db_schema_sqlite.sql. Designed to be called once per tag right after
extract_all() succeeds - either from the GUI (as an optional step) or
from a standalone bulk-import script.

No server needed - sqlite3 is part of Python's standard library. For
multiple people to accumulate into the SAME database over time, point
everyone's app at one shared database file (e.g. on a network drive)
rather than each person having their own local file.

Safe for repeated imports: re-importing the same (project, tag_no)
UPDATES that equipment_master row (via INSERT ... ON CONFLICT ... DO
UPDATE) and replaces its child rows (element_data, platform_data,
weight_summation) rather than accumulating duplicates.
"""
import sqlite3
import os


class DBConfig:
    """Just a database file path - no host/user/password needed."""
    def __init__(self, db_path="pvelite_data.db"):
        self.db_path = db_path

    def connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        return conn


def ensure_schema(config: DBConfig, schema_file="db_schema_sqlite.sql"):
    """
    Creates the database file and tables if they don't already exist.
    Safe to call every time the app starts.

    ALTER TABLE ADD COLUMN statements are run separately, one at a time,
    ignoring "duplicate column" errors - SQLite has no
    "ADD COLUMN IF NOT EXISTS", and executescript() would otherwise
    throw on the second run of the app once a column has already been
    added once (this was caught by testing a second run, not assumed).
    """
    conn = config.connect()
    try:
        with open(schema_file, "r") as f:
            full_sql = f.read()

        alter_statements = []
        other_statements = []
        for statement in full_sql.split(";"):
            # Strip SQL line-comments before checking/using the
            # statement - a comment directly above "ALTER TABLE" was
            # making the prefix check below see "--..." instead and
            # misclassify it, caught by testing a real multi-run
            # scenario rather than assumed correct.
            lines = [ln for ln in statement.splitlines() if not ln.strip().startswith("--")]
            cleaned = "\n".join(lines).strip()
            if not cleaned:
                continue
            if cleaned.upper().startswith("ALTER TABLE"):
                alter_statements.append(cleaned)
            else:
                other_statements.append(statement.strip())

        conn.executescript(";\n".join(other_statements) + ";")
        for stmt in alter_statements:
            try:
                conn.execute(stmt)
            except sqlite3.OperationalError as exc:
                if "duplicate column name" not in str(exc):
                    raise
        conn.commit()
    finally:
        conn.close()


def test_connection(config: DBConfig):
    """Returns (True, "") on success, (False, error_message) on failure -
    used by the GUI to validate the database file/path before an import."""
    try:
        conn = config.connect()
        conn.execute("SELECT 1")
        conn.close()
        return True, ""
    except sqlite3.Error as exc:
        return False, str(exc)


def _none_if_na(value):
    """weight_summation values can be the literal '...' from PV Elite's
    own table (not applicable) - NULL is the correct database
    representation of that, not the string itself."""
    if value is None or value == "...":
        return None
    return value


def write_result(config: DBConfig, project, category, tag_no, equipment_name,
                  extractor, result, source_file, imported_by=None):
    """
    Writes one tag's full result (equipment_master row + all child rows)
    into the database. `extractor` is the PVEliteDynamicExtractor
    instance (for file_kind); `result` is its ExtractionResult.
    """
    lw = result.loads_and_weights
    conn = config.connect()
    try:
        cur = conn.cursor()

        equipment_row = {
            "project": project, "category": category, "tag_no": tag_no,
            "equipment_name": equipment_name, "file_kind": extractor.file_kind,
            "warnings_count": len(result.warnings),
            "orientation": lw.get("orientation"), "support_type": lw.get("support_type"),
            "asme_code": lw.get("asme_code"), "division": lw.get("division"),
            "code_edition_year": lw.get("code_edition_year"), "mdmt_c": lw.get("mdmt_c"),
            "shell_moc": lw.get("shell_moc"), "shell_moc_class": lw.get("shell_moc_class"),
            "shell_uns": lw.get("shell_uns"),
            "tl_tl_mm": lw.get("tl_tl_mm"), "max_diameter_mm": lw.get("max_diameter_mm"),
            "min_diameter_mm": lw.get("min_diameter_mm"), "l_over_d_ratio": lw.get("l_over_d_ratio"),
            "max_design_pressure_mpa": lw.get("max_design_pressure_mpa"),
            "max_design_temp_c": lw.get("max_design_temp_c"),
            "wind_code": lw.get("wind_code"), "seismic_code": lw.get("seismic_code"),
            "basic_wind_speed_kmh": lw.get("basic_wind_speed_kmh"), "site_class": lw.get("site_class"),
            "seismic_ss": lw.get("seismic_ss"), "seismic_s1": lw.get("seismic_s1"),
            "seismic_sds": lw.get("seismic_sds"), "seismic_sd1": lw.get("seismic_sd1"),
            "seismic_response_coeff_cs": lw.get("seismic_response_coefficient_cs"),
            "max_wind_pressure_kgm2": lw.get("max_wind_pressure_kgm2"),
            "wind_shear_kgf": lw.get("wind_shear_kgf"), "eq_shear_kgf": lw.get("eq_shear_kgf"),
            "wind_moment_kgfm": lw.get("wind_moment_kgfm"), "eq_moment_kgfm": lw.get("eq_moment_kgfm"),
            "fabricated_mt": lw.get("fabricated_mt"), "shop_test_mt": lw.get("shop_test_mt"),
            "shipping_mt": lw.get("shipping_mt"), "erected_mt": lw.get("erected_mt"),
            "empty_mt": lw.get("empty_mt"), "operating_mt": lw.get("operating_mt"),
            "total_element_weight_mt": lw.get("total_element_weight_mt"),
            "total_surface_area_m2": lw.get("total_surface_area_m2"),
            "has_platform": 1 if lw.get("has_platform") == "Yes" else 0,
            "has_insulation": 1 if lw.get("has_insulation") == "Yes" else 0,
            "has_fp": 1 if lw.get("has_fp") == "Yes" else 0,
            "total_platform_area_m2": lw.get("total_platform_area_m2"),
            "insulation_area_m2": lw.get("insulation_area_m2") if lw.get("insulation_area_m2") != "N/A" else None,
            "painting_area_m2": lw.get("painting_area_m2"),
            "fp_area_m2": lw.get("fp_area_m2") if lw.get("fp_area_m2") != "N/A" else None,
            "basering_type": lw.get("basering_type"), "bolt_moc": lw.get("bolt_moc"),
            "bolt_nominal_dia_mm": lw.get("bolt_nominal_dia_mm"),
            "bolt_circle_dia_mm": lw.get("bolt_circle_dia_mm"), "bolt_qty": lw.get("bolt_qty"),
            "imported_by": imported_by, "source_file": source_file,
        }

        cols = list(equipment_row.keys())
        placeholders = ", ".join(["?"] * len(cols))
        update_clause = ", ".join(f"{c}=excluded.{c}" for c in cols if c not in ("project", "tag_no"))
        sql = (
            f"INSERT INTO equipment_master ({', '.join(cols)}) VALUES ({placeholders}) "
            f"ON CONFLICT(project, tag_no) DO UPDATE SET {update_clause}"
        )
        cur.execute(sql, list(equipment_row.values()))

        # Look up the id explicitly (SQLite's lastrowid isn't reliable
        # across an ON CONFLICT UPDATE path the way it is for a plain
        # INSERT), then clear and re-insert child rows so a re-import
        # can't leave stale rows behind if e.g. an element count changes.
        cur.execute("SELECT id FROM equipment_master WHERE project=? AND tag_no=?", (project, tag_no))
        equipment_id = cur.fetchone()[0]

        cur.execute("DELETE FROM element_data WHERE equipment_id = ?", (equipment_id,))
        cur.execute("DELETE FROM platform_data WHERE equipment_id = ?", (equipment_id,))
        cur.execute("DELETE FROM weight_summation WHERE equipment_id = ?", (equipment_id,))

        if result.success:
            elem_sql = (
                "INSERT INTO element_data (equipment_id, element_no, element_name, component, "
                "material, class, uns_number, diameter_m, element_length_mm, minimum_thickness_mm, "
                "nominal_thickness_mm, corrosion_allowance_mm, external_pressure_mpa, "
                "design_pressure_mpa, design_temperature_c, allowable_tensile_stress_mpa, "
                "allowable_compressive_stress_mpa, surface_area_m2, inside_surface_area_m2, "
                "weight_mt, total_ele_empty_wgt_mt) VALUES "
                "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
            )
            for _, row in result.dataframe.iterrows():
                cur.execute(elem_sql, (
                    equipment_id, row.get("Element No"), row.get("Element Name"), row.get("Component"),
                    row.get("Material"), row.get("Class"), row.get("UNS Number"), row.get("Diameter (m)"),
                    row.get("Element Length (mm)"), row.get("Minimum Thickness (mm)"),
                    row.get("Nominal Thickness (mm)"), row.get("Corrosion Allowance (mm)"),
                    row.get("External Pressure (MPa)"), row.get("Design Pressure (MPa)"),
                    row.get("Design Temperature (degC)"), row.get("Allowable Tensile Stress (MPa)"),
                    row.get("Allowable Compressive Stress (MPa)"), row.get("Surface Area (m2)"),
                    row.get("Inside Surface Area (m2)"), row.get("Weight (MT)"),
                    row.get("Total Ele. Empty Wgt. (MT)"),
                ))

        plat_sql = (
            "INSERT INTO platform_data (equipment_id, element_name, diameter_mm, detail_id, "
            "start_angle_deg, end_angle_deg, width_mm, length_mm, area_m2) "
            "VALUES (?,?,?,?,?,?,?,?,?)"
        )
        for p in result.platform_list:
            cur.execute(plat_sql, (
                equipment_id, p["element_name"], p["diameter_mm"], p["detail_id"],
                p["start_angle_deg"], p["end_angle_deg"], p["width_mm"], p["length_mm"], p["area_m2"],
            ))

        ws_sql = (
            "INSERT INTO weight_summation (equipment_id, component, fabricated_kg, shop_test_kg, "
            "shipping_kg, erected_kg, empty_kg, operating_kg) VALUES (?,?,?,?,?,?,?,?)"
        )
        for category_name, cond_values in result.weight_summation.items():
            cur.execute(ws_sql, (
                equipment_id, category_name,
                _none_if_na(cond_values.get("Fabricated")), _none_if_na(cond_values.get("Shop Test")),
                _none_if_na(cond_values.get("Shipping")), _none_if_na(cond_values.get("Erected")),
                _none_if_na(cond_values.get("Empty")), _none_if_na(cond_values.get("Operating")),
            ))

        # --- mechanical_details (EAV): max Nominal Thickness per
        # Component, rolled up from element_data - matches the
        # "display maximum value for each category" mechanical-summary
        # convention, and generalizes cleanly to Heat Exchangers later
        # (Tubesheet/Channel/Floating Head categories would land in the
        # same table with no schema change).
        cur.execute("DELETE FROM mechanical_details WHERE equipment_id = ?", (equipment_id,))
        if result.success:
            mech_sql = (
                "INSERT INTO mechanical_details (equipment_id, component, property, value, units) "
                "VALUES (?,?,?,?,?)"
            )
            thickness_by_component = {}
            for _, row in result.dataframe.iterrows():
                comp = row.get("Component")
                thk = row.get("Nominal Thickness (mm)")
                if comp and thk is not None:
                    thickness_by_component[comp] = max(thickness_by_component.get(comp, 0), thk)
            for comp, max_thk in thickness_by_component.items():
                cur.execute(mech_sql, (equipment_id, comp, "Max Nominal Thickness", max_thk, "mm"))

        # --- documents: one row per report section, the source text the
        # vector DB's embeddings get built from. embedded=0 flags rows
        # that still need to be pushed to the vector store (or re-pushed
        # if the section text changed on a re-import).
        cur.execute("DELETE FROM documents WHERE equipment_id = ?", (equipment_id,))
        try:
            from pvelite_dynamic_extractor import _split_toplevel_sections
            sections = _split_toplevel_sections(extractor.text)
            doc_sql = (
                "INSERT INTO documents (equipment_id, section_title, section_text, embedded) "
                "VALUES (?,?,?,0)"
            )
            for title, body in sections:
                cur.execute(doc_sql, (equipment_id, title, body))
        except Exception:
            pass  # documents are a bonus for vector search - never block the core DB write on this

        conn.commit()
        return equipment_id
    finally:
        conn.close()
