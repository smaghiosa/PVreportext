"""
Bulk-run PVEliteDynamicExtractor across every report in a folder and
consolidate into a single Master_Element_Data workbook.

Usage:
    python bulk_extract.py <reports_dir> <output_xlsx>
"""
import sys
import glob
import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from pvelite_dynamic_extractor import PVEliteDynamicExtractor


_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_THIN_BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

# Fixed reference encoding tables - covering commonly-used ASME pressure
# vessel plate materials and wind/seismic design codes generally, not
# just whatever happens to appear in the current batch of reports. Using
# a fixed table (rather than building codes from only what's present in
# a given run) keeps a material's or code's number STABLE across future
# runs, so historical training data stays comparable as new tags/reports
# are added later.
_MATERIAL_REFERENCE_CODES = {
    "SA-285": 1, "SA-299": 2, "SA-302": 3, "SA-455": 4,
    "SA-515": 10, "SA-516": 11, "SA-537": 12,
    "SA-517": 20, "SA-533": 21,
    "SA-203": 30, "SA-204": 31,
    "SA-387": 40,  # Cr-Mo alloy steel (Gr 11/12/22/91)
    "SA-353": 50, "SA-522": 51,  # 9% nickel, low temp
    "SA-240": 60,  # stainless plate (304/304L/316/316L/321 etc.)
    "SA-106": 70, "SA-333": 71,  # piping materials, occasionally used for small nozzles
    "SA-36": 80,  # structural (common basering/bolt material)
    "SA-193": 90, "SA-194": 91,  # bolting materials
}
_WIND_SEISMIC_CODE_REFERENCE = {
    "ASCE/SEI 7-05": 1, "ASCE/SEI 7-10": 2, "ASCE/SEI 7-16": 3, "ASCE/SEI 7-22": 4,
    "UBC 97": 10, "IBC 2015": 11, "IBC 2018": 12, "IBC 2021": 13,
    "IS 875": 20, "IS 1893": 21,
    "BS 6399": 30, "EN 1991-1-4": 31, "EN 1998-1": 32,
}


def _reference_code_for(value, reference_table):
    """Matches `value` against a reference table by prefix (so "SA-516
    70" matches the "SA-516" entry, "SA-387 11" matches "SA-387", etc.),
    falling back to a new stable code appended after the fixed range if
    genuinely not in the reference list, rather than dropping it to 0."""
    if not value:
        return 0
    for key, code in reference_table.items():
        if value.upper().startswith(key.upper()):
            return code
    return 900 + (abs(hash(value)) % 99)  # stable-ish fallback for anything not in the reference list


def _style_sheet(ws):
    """Applies a consistent, readable style to any worksheet: colored
    bold header row, thin borders on all populated cells, and column
    widths sized to fit their content."""
    if ws.max_row < 1:
        return
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            cell.border = _THIN_BORDER
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 3, 10), 42)
    ws.row_dimensions[1].height = 30

def main(
    reports_dir,
    out_path,
    selected_categories=None
):

    all_frames = []
    all_warnings = []
    equipment_stats = []
    bom_rows_all = []
    loads_rows = []
    work_volume_rows = []
    ml_summary_rows = []
    platform_list_rows = []

    if selected_categories:

        pdf_files = []

        for category in selected_categories:

            pdf_files.extend(
                glob.glob(
                    os.path.join(
                        reports_dir,
                        category,
                        "*.pdf"
                    )
                )
            )

    else:

        pdf_files = glob.glob(
            os.path.join(
                reports_dir,
                "**",
                "*.pdf"
            ),
            recursive=True
        )

    print(f"\nTotal PDFs Found: {len(pdf_files)}")

    for path in sorted(pdf_files):

        fname = os.path.basename(path)

        equipment_category = os.path.basename(
            os.path.dirname(path)
        )

        print(
            f"Processing {fname} "
            f"[{equipment_category}] ..."
        )

        try:

            ext = PVEliteDynamicExtractor(path)

            result = ext.extract_all()

        except Exception as exc:

            print(f"  FAILED: {exc}")

            all_warnings.append({
                "Equipment": fname,
                "Warning": f"EXTRACTION CRASHED: {exc}"
            })

            continue

        print(
            f"  File kind: {ext.file_kind}"
            f" | elements: {result.stats.get('total_elements')}"
            f" | warnings: {len(result.warnings)}"
        )

        if not result.success:

            all_warnings.append({
                "Equipment": ext.equipment_tag,
                "Warning":
                "Extraction produced 0 elements - table layout not recognized."
            })

            continue

        df_temp = result.dataframe.copy()

        df_temp.insert(
            0,
            "Equipment Category",
            equipment_category
        )

        all_frames.append(df_temp)

        for w in result.warnings:
            all_warnings.append({"Equipment": ext.equipment_tag, "Warning": w})
        lw = result.loads_and_weights
        dt = lw.get("detail_type_totals_kg", {})
        equipment_stats.append({
            "Equipment Category": equipment_category,
            "Tag No": ext.equipment_tag,
            "Name": ext.equipment_name,
            "File Kind": ext.file_kind,
            "Elements": result.stats.get("total_elements"),
            "Total Weight (MT)": round(result.stats.get("total_weight_mt", 0), 3),
            "Total Surface Area (m2)": round(result.stats.get("total_surface_area_m2", 0), 3),
            "Warnings": len(result.warnings),
        })
        loads_rows.append({
            "Tag No": ext.equipment_tag,
            "Name": ext.equipment_name,
            "ASME Code": lw.get("asme_code"),
            "Division": lw.get("division"),
            "Code Edition Year": lw.get("code_edition_year"),
            "MDMT (degC)": lw.get("mdmt_c"),
            "Wind Design Code": lw.get("wind_code"),
            "Seismic Design Code": lw.get("seismic_code"),
            "Basic Wind Speed (Km/hr)": lw.get("basic_wind_speed_kmh"),
            "Site Class": lw.get("site_class"),
            "Seismic Ss": lw.get("seismic_ss"),
            "Seismic S1": lw.get("seismic_s1"),
            "Seismic SDS": lw.get("seismic_sds"),
            "Seismic SD1": lw.get("seismic_sd1"),
            "Seismic Response Coeff. Cs": lw.get("seismic_response_coefficient_cs"),
            "Max Wind Pressure (Kgs/m2)": lw.get("max_wind_pressure_kgm2"),
            "Wind Shear (Kgf)": lw.get("wind_shear_kgf"),
            "Earthquake Shear (Kgf)": lw.get("eq_shear_kgf"),
            "Wind Moment (Kgf-m)": lw.get("wind_moment_kgfm"),
            "Earthquake Moment (Kgf-m)": lw.get("eq_moment_kgfm"),
            "Fabricated Weight (MT)": lw.get("fabricated_mt"),
            "Erected Weight (MT)": lw.get("erected_mt"),
            "Operating Weight (MT)": lw.get("operating_mt"),
            "Shop Test Weight (MT)": lw.get("shop_test_mt"),
            "Bolt MOC": lw.get("bolt_moc"),
            "Bolt Nominal Dia (mm)": lw.get("bolt_nominal_dia_mm"),
            "Bolt Circle Dia (mm)": lw.get("bolt_circle_dia_mm"),
            "Number of Bolts": lw.get("bolt_qty"),
        })
        dt = lw.get("detail_type_totals_kg", {})
        work_volume_rows.append({
            "Tag No": ext.equipment_tag,
            "Name": ext.equipment_name,
            "TL-TL (mm)": lw.get("tl_tl_mm"),
            "Max Diameter (mm)": lw.get("max_diameter_mm"),
            "Min Diameter (mm)": lw.get("min_diameter_mm"),
            "L/D Ratio": lw.get("l_over_d_ratio"),
            "Erection Weight (MT)": lw.get("erected_mt"),
            "Platform Qualified": lw.get("has_platform"),
            "Total Platform Area (m2)": lw.get("total_platform_area_m2"),
            "Insulation Qualified": lw.get("has_insulation"),
            "Insulation Area (m2)": lw.get("insulation_area_m2"),
            "Painting Area (m2)": lw.get("painting_area_m2"),
            "FP Qualified": lw.get("has_fp"),
            "FP Area (m2)": lw.get("fp_area_m2"),
        })

        for p in result.platform_list:
            platform_list_rows.append({
                "Tag No": ext.equipment_tag,
                "Element Name": p["element_name"],
                "Diameter (mm)": p["diameter_mm"],
                "Detail Type": "Platform",
                "Detail ID": p["detail_id"],
                "Start Angle (deg)": p["start_angle_deg"],
                "End Angle (deg)": p["end_angle_deg"],
                "Platform Width (mm)": p["width_mm"],
                "Platform Length (mm)": p["length_mm"],
                "Platform Area (m2)": p["area_m2"],
            })

        for category, cond_values in result.weight_summation.items():
            bom_rows_all.append({
                "Tag No": ext.equipment_tag,
                "Components": category,
                "Fabricated": cond_values.get("Fabricated"),
                "Shop Test": cond_values.get("Shop Test"),
                "Shipping": cond_values.get("Shipping"),
                "Erected": cond_values.get("Erected"),
                "Empty": cond_values.get("Empty"),
                "Operating": cond_values.get("Operating"),
            })

        shell_moc_combined = " ".join(
            filter(None, [lw.get("shell_moc"), lw.get("shell_moc_class")])
        ) or None
        ml_summary_rows.append({
            # --- INPUT FEATURES (known up-front for a new tag, before
            # any PV Elite run - dimensions, design basis, MOC, codes) ---
            "Equipment Category": equipment_category,
            "Tag No": ext.equipment_tag,
            "Name": ext.equipment_name,
            "Orientation": lw.get("orientation"),
            "Support Type": lw.get("support_type"),
            "ASME Code": lw.get("asme_code"),
            "Division": lw.get("division"),
            "MDMT (degC)": lw.get("mdmt_c"),
            "Shell MOC": shell_moc_combined,
            "TL-TL (mm)": lw.get("tl_tl_mm"),
            "Max Diameter (mm)": lw.get("max_diameter_mm"),
            "Min Diameter (mm)": lw.get("min_diameter_mm"),
            "L/D Ratio": lw.get("l_over_d_ratio"),
            "Max Design Pressure (MPa)": lw.get("max_design_pressure_mpa"),
            "Max Design Temp (degC)": lw.get("max_design_temp_c"),
            "Wind Design Code": lw.get("wind_code"),
            "Seismic Design Code": lw.get("seismic_code"),
            "Basic Wind Speed (Km/hr)": lw.get("basic_wind_speed_kmh"),
            "Seismic Response Coeff. Cs": lw.get("seismic_response_coefficient_cs"),
            "Max Wind Pressure (Kgs/m2)": lw.get("max_wind_pressure_kgm2"),
            "Bolt MOC": lw.get("bolt_moc"),
            "Bolt Nominal Dia (mm)": lw.get("bolt_nominal_dia_mm"),
            "Bolt Circle Dia (mm)": lw.get("bolt_circle_dia_mm"),
            "Number of Bolts": lw.get("bolt_qty"),
            "Platform Qualified": lw.get("has_platform"),
            "Insulation Qualified": lw.get("has_insulation"),
            "FP Qualified": lw.get("has_fp"),

            # --- PREDICTION TARGETS (what a new tag's weight/WV/load
            # figures need to be forecast, using the above as inputs) ---
            "Fabricated Weight (MT)": lw.get("fabricated_mt"),
            "Erected Weight (MT)": lw.get("erected_mt"),
            "Operating Weight (MT)": lw.get("operating_mt"),
            "Shop Test Weight (MT)": lw.get("shop_test_mt"),
            "Wind Shear (Kgf)": lw.get("wind_shear_kgf"),
            "Earthquake Shear (Kgf)": lw.get("eq_shear_kgf"),
            "Wind Moment (Kgf-m)": lw.get("wind_moment_kgfm"),
            "Earthquake Moment (Kgf-m)": lw.get("eq_moment_kgfm"),
            "Total Platform Area (m2)": lw.get("total_platform_area_m2"),
            "Insulation Area (m2)": lw.get("insulation_area_m2"),
            "Painting Area (m2)": lw.get("painting_area_m2"),
            "FP Area (m2)": lw.get("fp_area_m2"),

            # --- Derived ratios (diagnostic / auxiliary features) ---
            "Wind Shear / Surface Area": lw.get("wind_shear_to_surface_area"),
            "EQ Shear / Operating Weight": lw.get("eq_shear_to_operating_weight"),
            "Fab Weight per Length (MT/m)": lw.get("fab_weight_per_length_mt_per_m"),
            "Fab Weight per Footprint (MT/m3)": lw.get("fab_weight_per_footprint_mt_per_m3"),
            "Fab Weight per Surface Area (kg/m2)": lw.get("fab_weight_per_surface_area"),
        })

    if not all_frames:
        print("No reports extracted successfully.")
        return

    master_df = pd.concat(all_frames, ignore_index=True)

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Master_Element_Data"
    ws1.append(list(master_df.columns))
    for c in ws1[1]:
        c.font = Font(bold=True)
    for row in master_df.itertuples(index=False):
        ws1.append(list(row))
    ws1.freeze_panes = "A2"

    ws_lw = wb.create_sheet("Loads_Foundation_Support")
    if loads_rows:
        cols = list(loads_rows[0].keys())
        ws_lw.append(cols)
        for c in ws_lw[1]:
            c.font = Font(bold=True)
        for row in loads_rows:
            ws_lw.append([row[c] for c in cols])
    ws_lw.freeze_panes = "A2"

    ws_wv = wb.create_sheet("Work_Volume")
    if work_volume_rows:
        cols = list(work_volume_rows[0].keys())
        ws_wv.append(cols)
        for c in ws_wv[1]:
            c.font = Font(bold=True)
        for row in work_volume_rows:
            ws_wv.append([row[c] for c in cols])
    ws_wv.freeze_panes = "A2"

    ws_ml = wb.create_sheet("ML_Summary")
    if ml_summary_rows:
        # Numeric-encode every categorical/text column (label encoding)
        # so the sheet is directly usable as ML input without further
        # preprocessing. Encoding maps are built from the values actually
        # present in this run. Shell MOC, Bolt MOC, Wind Design Code, and
        # Seismic Design Code are encoded against the FIXED reference
        # tables above, so their codes stay stable and cover commonly-
        # used values even beyond what's in this particular batch of
        # reports. Orientation/Support Type (only 2-3 possible values,
        # no natural external reference) still use run-derived codes.
        # "ASME Code" isn't separately encoded - "Division" (already
        # numeric 1/2) captures the same information.
        categorical_cols = ["Orientation", "Support Type", "Shell MOC",
                             "Wind Design Code", "Seismic Design Code", "Bolt MOC"]
        _reference_backed_cols = {
            "Shell MOC": _MATERIAL_REFERENCE_CODES,
            "Bolt MOC": _MATERIAL_REFERENCE_CODES,
            "Wind Design Code": _WIND_SEISMIC_CODE_REFERENCE,
            "Seismic Design Code": _WIND_SEISMIC_CODE_REFERENCE,
        }
        qualifier_cols = ["Platform Qualified", "Insulation Qualified", "FP Qualified"]
        encoding_maps = {}
        for col in categorical_cols:
            values = sorted({row.get(col) for row in ml_summary_rows if row.get(col) is not None})
            if col in _reference_backed_cols:
                ref = _reference_backed_cols[col]
                encoding_maps[col] = {v: _reference_code_for(v, ref) for v in values}
            else:
                encoding_maps[col] = {v: i + 1 for i, v in enumerate(values)}  # 0 reserved for "unknown/missing"

        for row in ml_summary_rows:
            for col in categorical_cols:
                row[col + " Code"] = encoding_maps[col].get(row.get(col), 0)
            for col in qualifier_cols:
                row[col + " (1/0)"] = 1 if row.get(col) == "Yes" else 0

        # Rebuild column order explicitly so each "Code"/"(1/0)" column
        # sits right after its raw text counterpart, instead of every
        # encoded column landing at the end after the prediction targets
        # and ratios (which is where plain insertion order would put them).
        base_cols = list(ml_summary_rows[0].keys())
        generated_code_cols = {c + " Code" for c in categorical_cols}
        generated_qual_cols = {c + " (1/0)" for c in qualifier_cols}
        ordered_cols = []
        for c in base_cols:
            if c in generated_code_cols or c in generated_qual_cols:
                continue
            ordered_cols.append(c)
            if c in categorical_cols:
                ordered_cols.append(c + " Code")
            if c in qualifier_cols:
                ordered_cols.append(c + " (1/0)")
        cols = ordered_cols
        ws_ml.append(cols)
        for c in ws_ml[1]:
            c.font = Font(bold=True)
        for row in ml_summary_rows:
            ws_ml.append([row.get(c) for c in cols])

        # A lookup-key sheet so the numeric codes are traceable back to
        # their original text values, since the codes alone aren't
        # self-explanatory without this.
        ws_key = wb.create_sheet("ML_Encoding_Key")
        ws_key.append(["Column", "Text Value", "Numeric Code", "Source"])
        for c in ws_key[1]:
            c.font = Font(bold=True)
        for col, mapping in encoding_maps.items():
            for text_val, code in sorted(mapping.items(), key=lambda x: x[1]):
                ws_key.append([col, text_val, code, "This run"])
        # Also list the FULL fixed reference tables (materials and wind/
        # seismic codes), including entries not present in this run's
        # reports, so the key doubles as a general lookup for encoding
        # future tags/materials/codes consistently.
        for material, code in sorted(_MATERIAL_REFERENCE_CODES.items(), key=lambda x: x[1]):
            ws_key.append(["Material Reference (Shell/Bolt MOC)", material, code, "Fixed reference"])
        for code_name, code in sorted(_WIND_SEISMIC_CODE_REFERENCE.items(), key=lambda x: x[1]):
            ws_key.append(["Wind/Seismic Code Reference", code_name, code, "Fixed reference"])
        ws_key.freeze_panes = "A2"
    ws_ml.freeze_panes = "A2"

    ws_plat = wb.create_sheet("Platform_List")
    if platform_list_rows:
        cols = list(platform_list_rows[0].keys())
        ws_plat.append(cols)
        for c in ws_plat[1]:
            c.font = Font(bold=True)
        for row in platform_list_rows:
            ws_plat.append([row[c] for c in cols])
    ws_plat.freeze_panes = "A2"

    ws_bom = wb.create_sheet("weight_summ")
    if bom_rows_all:
        cols = ["Tag No", "Components", "Fabricated", "Shop Test",
                "Shipping", "Erected", "Empty", "Operating"]
        ws_bom.append(cols)
        for c in ws_bom[1]:
            c.font = Font(bold=True)
        for row in bom_rows_all:
            ws_bom.append([row.get(c) for c in cols])
    ws_bom.freeze_panes = "A2"

    ws2 = wb.create_sheet("Equipment_Summary")
    if equipment_stats:
        cols = list(equipment_stats[0].keys())
        ws2.append(cols)
        for c in ws2[1]:
            c.font = Font(bold=True)
        for row in equipment_stats:
            ws2.append([row[c] for c in cols])
    ws2.freeze_panes = "A2"

    if all_warnings:
        ws3 = wb.create_sheet("Warnings")
        ws3.append(["Equipment", "Warning"])
        for c in ws3[1]:
            c.font = Font(bold=True)
        for w in all_warnings:
            ws3.append([w["Equipment"], w["Warning"]])
        ws3.freeze_panes = "A2"

    for sheet_name in wb.sheetnames:
        _style_sheet(wb[sheet_name])

    wb.save(out_path)
    print(f"\nWrote {len(master_df)} element rows across {len(equipment_stats)} vessels -> {out_path}")
    print(f"Total warnings: {len(all_warnings)}")


if __name__ == "__main__":

    reports_dir = sys.argv[1] if len(sys.argv) > 1 else "Input"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "Master_Element_Data.xlsx"

    selected_categories = None

    folders = sorted([
        f for f in os.listdir(reports_dir)
        if os.path.isdir(os.path.join(reports_dir, f))
    ])

    print("\nAvailable Categories:")
    print("----------------------")

    for i, folder in enumerate(folders, start=1):
        print(f"{i}. {folder}")

    choice = input(
        "\nEnter folder numbers separated by comma "
        "(Press Enter for ALL): "
    ).strip()

    if choice:

        selected_categories = [
            folders[int(x.strip()) - 1]
            for x in choice.split(",")
        ]

    main(
        reports_dir,
        out_path,
        selected_categories
    )
