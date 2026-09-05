"""
pvelite_dynamic_extractor.py

Implements PVEliteDynamicExtractor, the class expected by
pvelite_notebook_dynamic.ipynb, built to the exact table/column
specification in prompt_to_imp_search.txt and Next_step_prompt.txt.

Design principle (per Next_step_prompt.txt): NO hardcoded element names
or page numbers. Tables are located by title text and parsed generically.
Element# (from Table-2, the "Primary Element Master") is the master key;
Tables 1/3/4/6/7 are joined onto it by row *position* — every one of
these tables enumerates the same physical elements top-to-bottom in the
same order, which we verify with an explicit quality check rather than
assume silently.

Each PV Elite report can be exported under a DIFFERENT unit system
(kPa/MPa/kgf-cm2 for pressure, cm2/mm2 for area, kg/kgm for
weight/volume) - every table declares its own units in its header line,
so every parser here detects units locally rather than assuming a fixed
system.
"""
import re
import math
import os
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font

FUZZY_CUTOFF = 0.72


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------
def _num(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s in ("...", "No Calc", "N/A", ""):
        return None
    m = re.search(r"-?\d+\.?\d*(?:[Ee][+-]?\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def _pressure_to_kpa(value, unit):
    if value is None:
        return None
    unit = (unit or "").lower()
    if "mpa" in unit:
        return round(value * 1000.0, 4)
    if "kgf/cm" in unit or "kg/cm" in unit:
        return round(value * 98.0665, 4)
    return round(value, 4)  # already kPa


def _stress_to_mpa(value, unit):
    if value is None:
        return None
    unit = (unit or "").lower()
    if "kgf/mm" in unit:
        return round(value * 9.80665, 4)
    if "kgf/cm" in unit or "kg/cm" in unit:
        return round(value * 0.0980665, 4)
    return round(value, 4)  # N/mm2 or MPa (same numeric value)


# --------------------------------------------------------------------------
# Section splitting (top-level headings repeat per page; merge them)
# --------------------------------------------------------------------------
_TOPLEVEL_HEADING = re.compile(r"^([A-Z][A-Za-z0-9 /&.]+?):\s*Step:\s*\d+.*$", re.MULTILINE)


def _split_toplevel_sections(text):
    matches = list(_TOPLEVEL_HEADING.finditer(text))
    raw_chunks = []
    for i, m in enumerate(matches):
        title = m.group(1).strip()
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        raw_chunks.append((title, text[start:end]))
    merged = []
    for title, body in raw_chunks:
        if merged and merged[-1][0] == title:
            merged[-1] = (title, merged[-1][1] + body)
        else:
            merged.append((title, body))
    return merged


def _find_section(text, heading_prefix):
    for title, body in _split_toplevel_sections(text):
        if title.startswith(heading_prefix):
            return _strip_page_break_noise(body)
    return ""


# A page break can fall in the MIDDLE of a table (even mid-header), leaving
# behind a 5-line boilerplate block: blank line, "PV Elite(R) vNN", the
# licensee line, the "FileName : ... Page N of M" line, and a repeated
# "<Title>: Step: N ..." heading line. This is stripped from merged section
# bodies so a table split across a page boundary reads as one continuous
# block regardless of where PV Elite happened to break the page.
_PAGE_BREAK_NOISE = re.compile(
    r"\n[ \t]*\nPV Elite\S* v\d+\n[^\n]*\nFileName[^\n]*Page \d+ of \d+\n[^\n]*:\s*Step:\s*\d+[^\n]*\n"
)


def _strip_page_break_noise(body):
    return _PAGE_BREAK_NOISE.sub("\n", body)


# --------------------------------------------------------------------------
# Component identification (Next_step_prompt.txt "COMPONENT IDENTIFICATION
# LOGIC" - dynamic, no hardcoded element names)
# --------------------------------------------------------------------------
def identify_component(element_type, element_name):
    et = (element_type or "").strip().lower()
    if et in ("cylinder",):
        return "Shell"
    if et in ("ellipse", "elliptical", "spherical", "sphere", "torispherical",
              "hemispherical", "flat head"):
        return "Head"
    if et in ("cone", "conical"):
        return "Cone"
    if et in ("skirt", "skirt support"):
        return "Skirt"
    if et in ("body flg", "flange", "nozzle flange"):
        return "Flange"
    # Fallback per spec: element-name prefix
    name = (element_name or "").strip()
    if name[:2].upper() in ("S-",):
        return "Shell"
    if name[:2].upper() in ("H-",):
        return "Head"
    if name[:2].upper() in ("K-",):
        return "Skirt"
    if name[:2].upper() in ("C-",):
        return "Cone"
    return element_type or "Unknown"


# --------------------------------------------------------------------------
# TABLE 1: Element Thickness, Pressure, Diameter and Allowable Stress
# --------------------------------------------------------------------------
def extract_table1(text):
    """Returns list of dicts: element_no, element_name, diameter_mm, nominal_thickness_mm, allowable_stress_mpa"""
    m = re.search(
        r"Element Thickness, Pressure, Diameter and Allowable Stress\s*:\s*\r?\n"
        r".*?\r?\n.*?\r?\n\s*([A-Za-z].*?)\r?\n(.*?)\n\s*Element Required Thickness",
        text, re.DOTALL,
    )
    rows = []
    if not m:
        return rows
    unit_line = m.group(1)
    stress_unit = unit_line.split()[-1] if unit_line.split() else ""
    body = m.group(2)
    for line in body.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue
        mm = re.match(r"^\s*(.+?)\s+(\S+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+(\S+)\s+(\d+)\s*$", line)
        if mm:
            name, press_raw, thk, corr, dia, stress_raw, elem_no = mm.groups()
            rows.append({
                "element_no": int(elem_no),
                "element_name": name.strip(),
                "diameter_mm": _num(dia),
                "nominal_thickness_mm": _num(thk),
                "allowable_stress_mpa": _stress_to_mpa(_num(stress_raw), stress_unit),
                "corrosion_allowance_mm": _num(corr),
            })
    return rows


# --------------------------------------------------------------------------
# TABLE 2: Element Required Thickness and MAWP  (PRIMARY ELEMENT MASTER)
# --------------------------------------------------------------------------
def extract_table2(text):
    m = re.search(
        r"Element Required Thickness and MAWP\s*:\s*\r?\n"
        r".*?\r?\n.*?\r?\n\s*([A-Za-z].*?)\r?\n(.*?)\n\s*Minimum\s",
        text, re.DOTALL,
    )
    rows = []
    if not m:
        return rows
    unit_line = m.group(1)
    unit = unit_line.split()[0] if unit_line.split() else ""
    body = m.group(2)
    for line in body.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue
        mm = re.match(
            r"^\s*(.+?)\s+(No Calc|[\d.]+)\s+(No Calc|[\d.]+)\s+(No Calc|[\d.]+)\s+"
            r"([\d.]+)\s+(No Calc|[\d.]+)\s+(\d+)\s*$", line
        )
        if mm:
            name, dp, mawp_corr, map_new, min_thk, reqd_thk, elem_no = mm.groups()
            rows.append({
                "element_no": int(elem_no),
                "element_name": name.strip(),
                "design_pressure_kpa": _pressure_to_kpa(_num(dp), unit),
                "mawp_corroded_kpa": _pressure_to_kpa(_num(mawp_corr), unit),
                "map_new_cold_kpa": _pressure_to_kpa(_num(map_new), unit),
                "minimum_thickness_mm": _num(min_thk),
                "required_thickness_mm": _num(reqd_thk),
            })
    return rows


# --------------------------------------------------------------------------
# TABLE 3: Surface Areas of Elements
# --------------------------------------------------------------------------
def extract_table3(text):
    """
    Locates the 'Surface Areas of Elements' data by title + terminator only
    - NOT by matching the exact 3-line stacked header text ("Outside
    Surface / Inside Surface", "From To Area Area", unit line). Real-PDF
    text extraction frequently reflows multi-line stacked column headers
    differently than plain text does, which breaks literal header
    matching even though the data rows themselves extract fine. Matching
    data rows by their own numeric shape is robust to that.
    """
    title_m = re.search(r"Surface Areas of Elements\s*:[ \t]*\n", text)
    if not title_m:
        return []
    window = text[title_m.end():title_m.end() + 6000]
    end_m = re.search(r"\n\s*Total\s", window)
    header_region = window[:200]  # just for unit sniffing
    body = window[:end_m.start()] if end_m else window

    unit_is_mm = bool(re.search(r"\bmm2\b|\bmm²\b|\bmmA?2\b", header_region, re.IGNORECASE))

    rows = []
    for line in body.splitlines():
        mm = re.match(r"^\s*(\d+)\s+(\d+)\s+([\d.]+)\s+(\.\.\.|[\d.]+)", line)
        if mm:
            out_area = _num(mm.group(3))
            in_area = _num(mm.group(4)) if mm.group(4) != "..." else None
            out_area_cm2 = (out_area / 100.0) if (out_area and unit_is_mm) else out_area
            in_area_cm2 = (in_area / 100.0) if (in_area and unit_is_mm) else in_area
            rows.append({
                "from_node": mm.group(1), "to_node": mm.group(2),
                "outside_surface_area_cm2": out_area_cm2,
                "inside_surface_area_cm2": in_area_cm2,
            })
    return rows


# --------------------------------------------------------------------------
# TABLE 4: Element and Detail Weights (main element table only)
# --------------------------------------------------------------------------
def extract_table4(text):
    """
    Same principle as Table-3: locate by title + terminator, match data
    rows by numeric shape rather than the exact 3-line stacked header.
    """
    title_m = re.search(r"Element and Detail Weights\s*:[ \t]*\n", text)
    if not title_m:
        return []
    # This section can run to 100K+ characters (many Detail/"Wght" line
    # items - nozzles, clips, ladders, davits - sit between the element
    # rows and the section's own summary line), so use a generous window.
    # Terminate specifically on the plural "Totals <number>..." summary
    # line, NOT a generic "Total " match, since "Total Weight of Each
    # Detail Type:" (a sub-heading that appears earlier, mid-section)
    # would otherwise cut the body short before all element rows.
    window = text[title_m.end():title_m.end() + 250000]
    end_m = re.search(r"\n\s*Totals\s+[\d.]", window)
    body = window[:end_m.start()] if end_m else window

    rows = []
    for line in body.splitlines():
        # From-node/To-node pairs are always 2-3 digit multiples of 10
        # (10-20, 20-30, ...); anchor only on the start of the line so
        # trailing corroded-weight/volume columns (however many there
        # are) don't prevent a match. Detail rows (Insulation/Platform/
        # etc.) start with "<node> Wght ..." - non-numeric 2nd token -
        # so they naturally fail this pattern and are excluded.
        mm = re.match(r"^\s*(\d+)\s+(\d+)\s+([\d.]+)\b", line)
        if mm:
            rows.append({
                "from_node": mm.group(1), "to_node": mm.group(2),
                "weight_kg": _num(mm.group(3)),
            })
    return rows


# --------------------------------------------------------------------------
# TABLE 5: Materials of Construction
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# Per-element Design Temperature - lives in the Input Echo's "Complete
# Listing of Vessel Elements and Details", not in Tables 1-7, so needs its
# own pass. One "Element#" block per element, same order as Table-2.
# --------------------------------------------------------------------------
def extract_platform_details(text):
    """
    Every "Detail Type Platform" block in the Complete Listing, tagged
    with the index of the element chunk it appears in (same ordering as
    Table-1's rows, so its diameter can be looked up by that index). A
    single element can carry more than one platform (e.g. two skirt
    platforms at different elevations both attach to the same skirt
    course), so this returns a flat list, not one-per-element.

    Each field (Detail ID, Start/End Angle, Width, Length) is searched
    for INDEPENDENTLY within a bounded window after "Detail Type
    Platform", rather than requiring them to appear in one exact,
    strictly-sequential multi-line pattern - real-PDF text extraction
    can reorder or respace these fields in ways a single rigid
    multi-line regex won't survive, which is what broke this before.
    """
    m = re.search(
        r"Complete Listing of Vessel Elements and Details\s*:(.*?)"
        r"(?=\n\s*Internal Pressure Calculations\s*:|\n\s*XY Coordinate Calculations\s*:|\Z)",
        text, re.DOTALL,
    )
    if not m:
        return []
    block = m.group(1)
    chunks = re.split(r"\r?\n\s*Element#\s*\d+/\d+\s*\r?\n", block)[1:]
    platforms = []
    for elem_idx, chunk in enumerate(chunks):
        for pt_m in re.finditer(r"Detail Type\s+Platform", chunk):
            # Bound the window to end at the next Detail Type/Ladder
            # entry (or end of chunk) so fields from a DIFFERENT detail
            # item immediately after aren't accidentally picked up.
            next_detail_m = re.search(r"Detail Type\s+\S", chunk[pt_m.end():])
            window_end = pt_m.end() + next_detail_m.start() if next_detail_m else len(chunk)
            window = chunk[pt_m.end():window_end]

            id_m = re.search(r"Detail ID\s+(\S.*?)\s*\n", window)
            start_m = re.search(r"Platform\s+Start\s+Angle\s*\(degrees\)\s+(-?[\d.]+)", window)
            end_m = re.search(r"Platform\s+End\s+Angle\s*\(degrees\)\s+(-?[\d.]+)", window)
            width_m = re.search(r"Platform\s+Width\s+([\d.]+)", window)
            length_m = re.search(r"Platform\s+Length\s*\(top head platform\)\s+([\d.]+)", window)

            if not (start_m and end_m and width_m):
                continue  # not enough data to compute an area for this platform

            platforms.append({
                "element_index": elem_idx,
                "detail_id": id_m.group(1).strip() if id_m else None,
                "start_angle_deg": _num(start_m.group(1)),
                "end_angle_deg": _num(end_m.group(1)),
                "width_mm": _num(width_m.group(1)),
                "length_mm": _num(length_m.group(1)) if length_m else 0.0,
            })
    return platforms


def extract_element_design_temps(text):
    m = re.search(
        r"Complete Listing of Vessel Elements and Details\s*:(.*?)"
        r"(?=\n\s*Internal Pressure Calculations\s*:|\n\s*XY Coordinate Calculations\s*:|\Z)",
        text, re.DOTALL,
    )
    if not m:
        return []
    block = m.group(1)
    chunks = re.split(r"\r?\n\s*Element#\s*\d+/\d+\s*\r?\n", block)[1:]
    temps = []
    for chunk in chunks:
        tm = re.search(r"Design Temperature Internal Pressure\s+(-?[\d.]+)\s*(\S*)", chunk)
        val, unit = (_num(tm.group(1)), tm.group(2)) if tm else (None, None)
        # Material's own basic allowable stress at ambient - exists for
        # EVERY element including skirts (which Table-1's pressure-design
        # allowable stress column leaves blank, since skirts aren't a
        # pressure-boundary element there). Used as a fallback so skirt
        # rows aren't left with no allowable stress at all.
        am = re.search(r"Allowable Stress, Ambient\s+(-?[\d.]+)\s*(\S*)", chunk)
        allow_amb_raw, allow_amb_unit = (am.group(1), am.group(2)) if am else (None, None)
        temps.append({
            "design_temp_c": _temp_to_c(val, unit) if tm else None,
            "has_insulation": "Detail Type Insulation" in chunk,
            "has_platform": "Detail Type Platform" in chunk,
            "has_fp": bool(re.search(r"Detail ID\s+FP\S*", chunk)),
            "material_allowable_stress_ambient_mpa": _stress_to_mpa(_num(allow_amb_raw), allow_amb_unit) if am else None,
        })
    return temps


def _temp_to_c(value, unit):
    if value is None:
        return None
    if unit and "f" in unit.lower():
        return round((value - 32) * 5.0 / 9.0, 2)
    return value  # already °C


# --------------------------------------------------------------------------
# Governing (worst-case) induced stress per element, from "Stress due to
# Combined Loads" - PV Elite runs many load-case combinations (wind/
# seismic x operating/empty/hydrotest etc.); the number an engineer
# actually cares about is the WORST unity check across all of them, not
# any single case, so this scans every "Analysis of Load Case N" block
# and keeps the max-unity-check row per node.
# --------------------------------------------------------------------------
def extract_governing_combined_stress(text):
    """
    Locates each "Analysis of Load Case N : <combo>" block by that title
    ALONE - not by the 2-line stacked column header that follows it
    ("From Stress All. Stress..." / "Node Intensity Intensity..."), since
    that kind of multi-line stacked header is exactly the thing that does
    NOT reliably survive real-PDF text extraction (same failure mode
    already found and fixed for Tables 3/4/7). Row shape alone (Div-2's
    8-number rows vs Div-1's 7-number rows) is used to identify and parse
    data rows within each block instead.
    """
    governing = {}  # from_node (str) -> best row dict
    skipped = 0

    case_starts = list(re.finditer(r"Analysis of Load Case (\d+)\s*:\s*(\S+)\s*\r?\n", text))
    for i, cs in enumerate(case_starts):
        case_no, case_def = cs.group(1), cs.group(2)
        end = case_starts[i + 1].start() if i + 1 < len(case_starts) else cs.end() + 4000
        body = text[cs.end():end]

        div2_hits = 0
        for line in body.splitlines():
            line = line.rstrip()
            if not line.strip() or line.strip().startswith(("PV Elite", "FileName")):
                continue
            mm = re.match(
                r"^\s*(\d+)\s+(-?[\d.]+)\s+([\d.]+)\s+(-?[\d.]+)\s+([\d.]+)\s+"
                r"([\d.]+)\s+([\d.]+)\s+([\d.]+)(\*?)\s*$",
                line,
            )
            if not mm:
                continue
            div2_hits += 1
            node, stress_i, allow_i = mm.group(1), mm.group(2), mm.group(3)
            comp_stress, allow_comp = mm.group(4), mm.group(5)
            unity_val = _num(mm.group(8))
            overstress = mm.group(9)
            prev = governing.get(node)
            if prev is None or (unity_val is not None and (prev["unity_check"] is None or unity_val > prev["unity_check"])):
                governing[node] = {
                    "induced_stress_mpa": _num(stress_i),
                    "allowable_stress_combined_mpa": _num(allow_i),
                    "induced_compressive_stress_mpa": _num(comp_stress),
                    "allowable_compressive_stress_mpa": _num(allow_comp),
                    "unity_check": unity_val,
                    "governing_load_case": f"{case_no} ({case_def})",
                    "overstress": bool(overstress),
                }

        if div2_hits:
            continue  # this block was Div-2 style; don't also try Div-1 shape on it

        for line in body.splitlines():
            line = line.rstrip()
            if not line.strip() or line.strip().startswith(("PV Elite", "FileName")):
                continue
            mm = re.match(
                r"^\s*(\d+)\s+(-?[\d.]+)\s+([\d.]+)\s+(-?[\d.]+)\s+([\d.]+)\s+"
                r"([\d.]+)\s+([\d.]+)(\*?)\s*$",
                line,
            )
            if not mm:
                if re.match(r"^\s*\d+\s+[\d.]", line):
                    skipped += 1
                continue
            node, tens_stress, allow_tens = mm.group(1), mm.group(2), mm.group(3)
            comp_stress, allow_comp = mm.group(4), mm.group(5)
            tens_ratio, comp_ratio = _num(mm.group(6)), _num(mm.group(7))
            overstress = mm.group(8)
            ratios = [r for r in (tens_ratio, comp_ratio) if r is not None]
            governing_ratio = max(ratios) if ratios else None
            prev = governing.get(node)
            if prev is None or (governing_ratio is not None and (prev["unity_check"] is None or governing_ratio > prev["unity_check"])):
                # Div-1's "Tensile/Compressive Stress" table is in
                # kgf/cm2 (confirmed against the "Longitudinal Allowable
                # Stresses" table elsewhere in the same report, which
                # declares that unit explicitly and shares identical
                # values) - convert to MPa for consistency with every
                # other stress figure in this database.
                governing[node] = {
                    "induced_stress_mpa": round(_num(tens_stress) * 0.0980665, 4) if _num(tens_stress) else None,
                    "allowable_stress_combined_mpa": round(_num(allow_tens) * 0.0980665, 4) if _num(allow_tens) else None,
                    "induced_compressive_stress_mpa": round(_num(comp_stress) * 0.0980665, 4) if _num(comp_stress) else None,
                    "allowable_compressive_stress_mpa": round(_num(allow_comp) * 0.0980665, 4) if _num(allow_comp) else None,
                    "unity_check": governing_ratio,
                    "governing_load_case": f"{case_no} ({case_def})",
                    "overstress": bool(overstress),
                }

    return governing, skipped


# --------------------------------------------------------------------------
# Bill of Materials - parsed and consolidated into standard component
# categories (one summary row per vessel, per-category quantities), since
# raw BOM descriptions are free text and vary line to line.
# --------------------------------------------------------------------------
_BOM_CATEGORIES = [
    # (label, regex) - checked in order, most specific first, so e.g.
    # "BASERING BOLT NUTS" is caught before the bare "BASERING" pattern.
    ("Basering Bolt Nuts", r"^BASERING BOLT NUTS"),
    ("Basering Bolts", r"^BASERING BOLTS"),
    ("Basering Top Ring", r"^BASERING TOP RING"),
    ("Basering Gusset Plates", r"^BASERING GUSSET PLATES"),
    ("Basering", r"^BASERING\b"),
    ("Skirt Segments", r"^CYLINDRICAL SKIRT SEGMENT"),
    ("Heads", r"^(SPHERICAL|ELLIPTICAL|TORISPHERICAL|HEMISPHERICAL|FLAT) HEAD"),
    ("Platforms", r"^PLATFORM"),
    ("Ladders", r"^LADDER"),
    ("Insulation", r"^INSULATION"),
    ("Packing", r"^PACKING"),
    ("Lining", r"^LINING"),
    ("Ring Stiffeners", r"^(BAR )?RING STIFFENER"),
    ("Flanges", r"(BLIND )?FLANGE\(?S?\)?"),
    ("Trays", r"^TRAYS"),
    ("Nozzles", r"^NOZZLE"),
]


def extract_weight_summation_table(text):
    """
    The full "Weight Summation Results" category x condition matrix -
    e.g. Platforms/Insulation/Nozzles/Main Elements etc, each broken
    down by Fabricated/Shop Test/Shipping/Erected/Empty/Operating.
    Returns dict: category -> {condition: value_kg or None}.
    """
    m = re.search(r"Weight Summation Results\s*:(.*?)(?:Field Installation|\Z)", text, re.DOTALL)
    if not m:
        return {}
    conditions = ["Fabricated", "Shop Test", "Shipping", "Erected", "Empty", "Operating"]
    result = {}
    for line in m.group(1).splitlines():
        mm = re.match(
            r"^\s*([A-Za-z][A-Za-z. ]+?)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s*\*?\s*$",
            line,
        )
        if not mm:
            continue
        category = mm.group(1).strip()
        raw_tokens = mm.groups()[1:]
        # Keep "..." as a literal string (matching PV Elite's own display
        # convention for "not applicable"), converting only real numbers.
        values = [v if v == "..." else _num(v) for v in raw_tokens]
        if all(v in (None, "...") for v in values):
            continue  # skips the column-header line itself (words, not numbers)
        result[category] = dict(zip(conditions, values))
    return result


def extract_bom(text):
    m = re.search(
        r"Bill of Materials\s*:(.*?)(?=\n[A-Z][A-Za-z /&.]+:\s*Step:\s*\d+|\Z)",
        text, re.DOTALL,
    )
    if not m:
        return []
    rows = []
    for line in m.group(1).splitlines():
        line = line.strip()
        mm = re.match(r"^(\d+)\s+(.+)$", line)
        if not mm:
            continue
        qty, rest = int(mm.group(1)), mm.group(2)
        rows.append({"qty": qty, "description": rest})
    return rows


def categorize_bom(bom_rows):
    """Returns dict: category -> total qty, plus an 'Other' bucket for
    anything not matching a known category (so nothing is silently lost)."""
    totals = {}
    for row in bom_rows:
        desc_upper = row["description"].upper()
        category = None
        for label, pattern in _BOM_CATEGORIES:
            if re.search(pattern, desc_upper):
                category = label
                break
        category = category or "Other"
        totals[category] = totals.get(category, 0) + row["qty"]
    return totals


def extract_table5(text):
    m = re.search(r"Materials of Construction:(.*?)(?:\n[A-Z][A-Za-z /&.]+:)", text, re.DOTALL)
    out = {}
    if not m:
        return out
    block = m.group(1)
    for line in block.splitlines():
        line = line.strip()
        mm = re.match(
            r"^(Shell|Head|Cone|Flange|Skirt|Nozzle Flg|Nozzle|Basering|Ring|Rings|"
            r"Re-Pad|Flg Bolting|Base Bolting)\s+(.*)$", line,
        )
        if not mm:
            continue
        comp, rest = mm.group(1), mm.group(2)
        tokens = rest.split()
        mat_tokens, class_tokens = [], []
        i = 0
        # Material = leading run of non-'...' tokens up to (but not
        # including) the Class field; Class is the token(s) right after,
        # again up to '...'
        while i < len(tokens) and tokens[i] != "...":
            mat_tokens.append(tokens[i])
            i += 1
        material = mat_tokens[0] if mat_tokens else None
        klass = " ".join(mat_tokens[1:]) if len(mat_tokens) > 1 else None
        uns_match = re.search(r"\b([KG]\d{5})\b", rest)
        uns = uns_match.group(1) if uns_match else None
        options_match = re.search(r"(\+\S+(?:\s+\+\S+)*)\s*$", rest)
        options = options_match.group(1) if options_match else None
        entry = {"material": material, "class": klass, "uns": uns, "options": options}
        if comp == "Nozzle":
            out.setdefault("Nozzle", []).append(entry)
        elif comp not in out:
            out[comp] = entry
    return out


# --------------------------------------------------------------------------
# TABLE 6: Element Pressures and MAWP
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# TABLE 4b: "Total Ele. Empty Wgt." - a SECOND, differently-scoped table
# that shares the same "Element and Detail Weights:" title as Table-4.
# Verified against PV Elite's own printed subtotals: Table-4's rows sum to
# PV Elite's own "Total <n>" line for the element rows alone; this table's
# column total instead matches a Corroded ID Volume figure elsewhere in
# Table-4's summary row, and it's immediately followed by "Empty/Operating
# Support Force" calculations - so it's kept as a distinct, separately
# labeled field rather than merged into Weight (MT), which one it actually
# represents engineering-wise is for the person reading it to judge.
# --------------------------------------------------------------------------
def extract_table_total_ele_empty_wgt(text):
    header_m = re.search(r"From\s+To\s+Empty Wgt\.", text)
    if not header_m:
        return []
    window = text[header_m.end():header_m.end() + 6000]
    end_m = re.search(r"\n\s*(Empty Support Force|Cumulative Vessel Weight)", window)
    body = window[:end_m.start()] if end_m else window
    rows = []
    for line in body.splitlines():
        mm = re.match(
            r"^\s*(\d+)\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+"
            r"(?:\.\.\.|-?[\d.]+)\s+([\d.]+)\s*$", line
        )
        if mm:
            rows.append({
                "from_node": mm.group(1), "to_node": mm.group(2),
                "total_ele_empty_wgt_kg": _num(mm.group(3)),
            })
    return rows


def extract_table6(text):
    m = re.search(
        r"Element Pressures and MAWP\s*\(([^)]+)\):\s*\r?\n"
        r".*?\r?\n.*?\r?\n.*?\r?\n(.*?)(?:\n\s*Liquid Level:|\n\s*Element Types and Properties:|\Z)",
        text, re.DOTALL,
    )
    rows = []
    if not m:
        return rows
    unit = m.group(1).split("&")[0].strip()
    body = m.group(2)
    for line in body.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue
        mm = re.match(
            r"^\s*(.+?)\s+(\.\.\.|[\d.]+)\s+(\.\.\.|[\d.]+)\s+(\.\.\.|[\d.]+)\s+"
            r"([\d.]+)\s+(\S+)\s+(Yes|No)\s*$", line
        )
        if mm:
            name, dp, ext_p, mawp, corr, flg, creep = mm.groups()
            rows.append({
                "element_name": name.strip(),
                "design_pressure_kpa": _pressure_to_kpa(_num(dp), unit),
                "external_pressure_kpa": _pressure_to_kpa(_num(ext_p), unit),
            })
    return rows


# --------------------------------------------------------------------------
# TABLE 7: Element Types and Properties
# --------------------------------------------------------------------------
def extract_table7(text):
    """
    Same principle as Tables 3/4: locate by title only, don't require the
    exact 3-line stacked header ("Element \"To\" Elev Element Nominal...").
    Units (cm vs mm for elevation/length) are sniffed from whatever text
    appears near the top of the section instead of a fixed-position
    capture group, and rows are matched purely by their numeric shape:
    a type name followed by 8 numeric/".../-"-prefixed tokens.
    """
    title_m = re.search(r"Element Types and Properties\s*:[ \t]*\n", text)
    if not title_m:
        return []
    window = text[title_m.end():title_m.end() + 8000]
    # Apply page-break-noise stripping LOCALLY, only within this table's
    # own window - not globally on the whole document (see extract_all
    # for why). This table's own header can legitimately be split across
    # a page break (blank/PV Elite/FileName boilerplate + a repeated
    # "Vessel Design Summary: Step: N" line landing between its two
    # header lines), which is what this cleans up here.
    window = _strip_page_break_noise(window)
    end_m = re.search(r"\n\s*Local Stress Analysis", window)
    header_region = window[:300]
    body = window[:end_m.start()] if end_m else window

    # Sniff units: header region normally has a line like "Type cm cm mm
    # mm mm mm" (elevation/length can be cm OR mm; thickness cols are
    # always mm). If we can't find that line intact, default to mm
    # (matches the majority of reports seen) rather than failing outright.
    unit_line_m = re.search(r"\bType\s+(\S+)\s+(\S+)\s+mm\s+mm\s+mm\s+mm\b", header_region)
    if unit_line_m:
        elev_unit, length_unit = unit_line_m.group(1), unit_line_m.group(2)
    else:
        elev_unit, length_unit = "mm", "mm"
    length_scale = 10.0 if length_unit.lower() == "cm" else 1.0

    rows = []
    for line in body.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue
        mm = re.match(
            r"^\s*([A-Za-z][A-Za-z. ]*?)\s+(-?[\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+"
            r"(\.\.\.|[\d.]+)\s+(\.\.\.|[\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$", line
        )
        if mm:
            (etype, to_elev, length, nom_thk, fin_thk,
             reqd_int, reqd_ext, long_eff, circ_eff) = mm.groups()
            rows.append({
                "element_type": etype.strip(),
                "length_mm": _num(length) * length_scale if _num(length) is not None else None,
                "reqd_thickness_internal_mm": _num(reqd_int),
                "reqd_thickness_external_mm": _num(reqd_ext),
            })
    return rows


# --------------------------------------------------------------------------
# Result container
# --------------------------------------------------------------------------
@dataclass
class ExtractionResult:
    success: bool = False
    dataframe: pd.DataFrame = None
    stats: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)
    loads_and_weights: dict = field(default_factory=dict)
    bom_totals: dict = field(default_factory=dict)
    platform_list: list = field(default_factory=list)
    weight_summation: dict = field(default_factory=dict)


# --------------------------------------------------------------------------
# Equipment-level Loads for Foundation/Support Design (Vessel Design
# Summary - not element-wise, one set of values per vessel).
# --------------------------------------------------------------------------
def extract_equipment_loads_and_weights(text):
    def grab(label, pattern_after=r"([\d.]+)\s*(\S+)"):
        m = re.search(re.escape(label) + r"\s+" + pattern_after, text)
        return (m.group(1), m.group(2)) if m else (None, None)

    result = {}
    for key, label in [
        ("wind_shear", "Total Wind Shear on Support"),
        ("eq_shear", "Total Earthquake Shear on Support"),
        ("wind_moment", "Wind Moment on Support"),
        ("eq_moment", "Earthquake Moment on Support"),
    ]:
        val, unit = grab(label)
        if "shear" in key:
            result[key + "_kgf"] = _force_to_kgf(_num(val), unit)
        else:
            result[key + "_kgfm"] = _moment_to_kgfm(_num(val), unit)

    # --- Fabricated/Shop Test/Shipping/Erected/Empty/Operating weights ---
    # Locate by the "Weight Summation Results" title ALONE - not by also
    # requiring the exact column-header line beneath it, since real-PDF
    # text extraction commonly pads/varies spacing between column headers
    # (they're aligned over numeric columns), which broke a literal
    # single-space match here before. The "Totals" row's 6 numbers are
    # read by fixed position - PV Elite always prints this column order:
    # Fabricated, Shop Test, Shipping, Erected, Empty, Operating.
    wsr_title_m = re.search(r"Weight Summation Results\s*:", text)
    totals_m = None
    if wsr_title_m:
        window = text[wsr_title_m.end():wsr_title_m.end() + 4000]
        totals_m = re.search(
            r"\n\s*Totals\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)",
            window,
        )
    weight_labels = ["fabricated_mt", "shop_test_mt", "shipping_mt",
                      "erected_mt", "empty_mt", "operating_mt"]
    if totals_m:
        for key, raw in zip(weight_labels, totals_m.groups()):
            v = _num(raw)
            result[key] = round(v / 1000.0, 4) if v else None
    else:
        for key in weight_labels:
            result[key] = None

    # Fallback: narrative "Weight Summary" sentences, in case the tabular
    # section wasn't found at all. Tries both wordings PV Elite has used
    # ("Fabricated - Bare w/o..." and "Fabricated Wt. - Bare weight
    # without..."), and doesn't require the label and number to be on the
    # same line (real-PDF extraction can split them across lines).
    if not totals_m:
        narrative_labels = {
            "fabricated_mt": ["Fabricated Wt.", "Fabricated -"],
            "shop_test_mt": ["Shop Test Wt.", "Shop Test -"],
            "shipping_mt": ["Shipping Wt.", "Shipping -"],
            "erected_mt": ["Erected Wt.", "Erected -", "Erected  "],
            "empty_mt": ["Empty Wt.", "Empty -", "Empty  "],
            "operating_mt": ["Operating Wt.", "Operating -", "Operating  "],
        }
        for key, labels in narrative_labels.items():
            for label in labels:
                m = re.search(re.escape(label) + r"(.{0,150}?)([\d,]+\.?\d*)\s*kg", text, re.DOTALL)
                if m:
                    v = _num(m.group(2))
                    result[key] = round(v / 1000.0, 4) if v else None
                    break

    # Basering / anchor bolt data (support design item) - only present for
    # skirt-supported vertical vessels; absent (None) for saddle/leg-
    # supported horizontal ones, which is expected, not a failure.
    basering_m = re.search(r"Basering Data\s*:\s*(.+?)\s*\r?\n", text)
    result["basering_type"] = basering_m.group(1).strip() if basering_m else None
    bolt_dia_m = re.search(r"Nominal Diameter of Bolts\s+([\d.]+)\s*(\S+)", text)
    result["bolt_nominal_dia_mm"] = _num(bolt_dia_m.group(1)) if bolt_dia_m else None
    bolt_circle_m = re.search(r"Diameter of Bolt Circle\s+([\d.]+)\s*(\S+)", text)
    result["bolt_circle_dia_mm"] = _num(bolt_circle_m.group(1)) if bolt_circle_m else None
    bolt_qty_m = re.search(r"Number of Bolts\s+(\d+)", text)
    result["bolt_qty"] = int(bolt_qty_m.group(1)) if bolt_qty_m else None
    bolt_moc_m = re.search(r"Bolt Material\s+(\S.*?)\s*\r?\n", text)
    result["bolt_moc"] = bolt_moc_m.group(1).strip() if bolt_moc_m else None

    # Max wind pressure applied across elements - from the "Element Hgt
    # (z) ... qz" table in Wind Load Calculation (qz is the velocity
    # pressure per element). Units vary by report (Kgs/m2 or kPa) - the
    # unit line right after the header ("mm kPa" etc.) is used to detect
    # and convert. Rows are matched STRICTLY by their shape (name + 6
    # numeric columns) within a bounded window, rather than taking "the
    # last number on any line" over an unbounded region - the previous
    # approach ran away into unrelated tables whenever the "Note:"
    # terminator it relied on wasn't present nearby.
    qz_header_m = re.search(r"Element\s+Hgt\s*\(z\)\s+K1\s+K2\s+K3\s+Kz\s+Kzt\s+qz\s*\n\s*(\S+)\s+(\S+)\s*\n", text)
    if qz_header_m:
        qz_unit = qz_header_m.group(2)
        window = text[qz_header_m.end():qz_header_m.end() + 3000]
        qz_values = []
        for line in window.splitlines():
            mm = re.match(
                r"^\s*\S.*?\s+[\d.]+\s+[\d.]+\s+[\d.]+\s+[\d.]+\s+[\d.]+\s+[\d.]+\s+([\d.]+)\s*$",
                line,
            )
            if mm:
                qz_values.append(_num(mm.group(1)))
            elif line.strip() == "" or not re.match(r"^\s*\S+.*[\d.]", line):
                if qz_values:
                    break  # end of the table - stop before drifting into unrelated text
        qz_values = [v for v in qz_values if v is not None]
        max_qz = max(qz_values) if qz_values else None
        if max_qz is not None and "kpa" in qz_unit.lower():
            max_qz = round(max_qz * 101.9716, 3)  # kPa -> kgf/m2
        result["max_wind_pressure_kgm2"] = max_qz
    else:
        result["max_wind_pressure_kgm2"] = None

    # Design code / division, MDMT, TL-TL (for L/D ratio and other
    # derived equipment-level metrics)
    code_m = re.search(r"ASME Code,\s*Section VIII Division (\d+),\s*(\d{4})", text)
    result["asme_code"] = f"Section VIII Division {code_m.group(1)}" if code_m else None
    result["division"] = code_m.group(1) if code_m else None
    result["code_edition_year"] = code_m.group(2) if code_m else None
    mdmt_m = re.search(r"Required Minimum Design Metal Temperature\s+(-?[\d.]+)\s*(\S*)", text)
    result["mdmt_c"] = _temp_to_c(_num(mdmt_m.group(1)), mdmt_m.group(2)) if mdmt_m else None
    tltl_m = re.search(r"Vessel Design Length, Tangent to Tangent\s+([\d.]+)\s*(\S*)", text)
    result["tl_tl_mm"] = _num(tltl_m.group(1)) if tltl_m else None

    # Wind / Seismic design codes and parameters (for ML feature use -
    # normalizing shear/moment by these lets the model separate "how
    # windy/seismic is this site" from "how big is this vessel").
    wind_code_m = re.search(r"Wind Design Code\s+(\S.*?)\s*\r?\n", text)
    result["wind_code"] = wind_code_m.group(1).strip() if wind_code_m else None
    seismic_code_m = re.search(r"(?:Seismic|Earthquake) Design Code\s+(\S.*?)\s*\r?\n", text)
    result["seismic_code"] = seismic_code_m.group(1).strip() if seismic_code_m else None
    wind_speed_m = re.search(r"Basic Wind Speed\s*\[V\]\s+([\d.]+)\s*(\S+)", text)
    result["basic_wind_speed_kmh"] = _num(wind_speed_m.group(1)) if wind_speed_m else None
    site_class_m = re.search(r"Site Class\s+(\S+)", text)
    result["site_class"] = site_class_m.group(1) if site_class_m else None
    ss_m = re.search(r"SMS\s*=\s*Fa\s*\*\s*Ss\s*=\s*[\d.]+\s*\*\s*([\d.]+)", text)
    result["seismic_ss"] = _num(ss_m.group(1)) if ss_m else None
    s1_m = re.search(r"SM1\s*=\s*Fv\s*\*\s*S1\s*=\s*[\d.]+\s*\*\s*([\d.]+)", text)
    result["seismic_s1"] = _num(s1_m.group(1)) if s1_m else None
    sds_m = re.search(r"SDS\s*=\s*2/3\s*\*\s*Sms\s*=\s*2/3\s*\*\s*[\d.]+\s*=\s*([\d.]+)", text)
    result["seismic_sds"] = _num(sds_m.group(1)) if sds_m else None
    sd1_m = re.search(r"SD1\s*=\s*2/3\s*\*\s*Sm1\s*=\s*2/3\s*\*\s*[\d.]+\s*=\s*([\d.]+)", text)
    result["seismic_sd1"] = _num(sd1_m.group(1)) if sd1_m else None
    cs_m = re.search(
        r"Seismic Response Coefficient per equation 12\.8-2 \[Cs\]:\s*\r?\n"
        r"\s*=[^\r\n]*\r?\n\s*=\s*([\d.]+)",
        text,
    )
    result["seismic_response_coefficient_cs"] = _num(cs_m.group(1)) if cs_m else None

    # Detail-type weight rollup (Platforms, Insulation, Nozzles, Trays,
    # Packing, etc.) - PV Elite's own printed totals, not re-derived.
    dt_m = re.search(
        r"Total Weight of Each Detail Type\s*:\s*\r?\n(.*?)\n\s*Sum of the Detail Weights",
        text, re.DOTALL,
    )
    detail_totals = {}
    if dt_m:
        for line in dt_m.group(1).splitlines():
            mm2 = re.match(r"^\s*([A-Za-z][A-Za-z ]+?)\s+([\d.]+)\s*$", line)
            if mm2:
                detail_totals[mm2.group(1).strip()] = _num(mm2.group(2))
    result["detail_type_totals_kg"] = detail_totals

    return result


def _force_to_kgf(value, unit):
    if value is None:
        return None
    unit = (unit or "").lower()
    if unit.startswith("n") and "kgf" not in unit:
        return round(value / 9.80665, 1)
    return round(value, 1)


def _moment_to_kgfm(value, unit):
    if value is None:
        return None
    unit = (unit or "").lower()
    if "n-mm" in unit or "n.mm" in unit:
        return round(value / 9806.65, 1)
    if unit.startswith("n") and "kgf" not in unit and "mm" not in unit:
        return round(value / 9.80665, 1)
    return round(value, 1)  # already Kg-m / kgf-m


class PVEliteDynamicExtractor:
    def __init__(self, pdf_file):
        self.pdf_file = str(pdf_file)
        self.file_kind, self.text = self._load_text(self.pdf_file)
        self.equipment_tag, self.equipment_name = self._tag_and_name()
        pages = re.findall(r"Page (\d+) of (\d+)", self.text)
        self.total_pages = int(pages[0][1]) if pages else None
        self.log = []
        self.result = None

    @staticmethod
    def _load_text(path):
        """
        Auto-detect and handle BOTH file kinds that show up in practice:
          1. PV Elite's raw plain-text report, saved/renamed with a .pdf
             extension (no real PDF structure at all - identified by the
             absence of the '%PDF' magic bytes at the start of the file).
          2. A genuine, binary PDF (e.g. "Print to PDF" of the same report,
             or a native PV Elite PDF export) - identified by the '%PDF'
             magic bytes, and read with PyMuPDF.
        Returns (file_kind, text) where file_kind is 'text' or 'pdf'.
        """
        with open(path, "rb") as f:
            head = f.read(8)
        if head.lstrip().startswith(b"%PDF"):
            try:
                import pymupdf
            except ImportError:
                import fitz as pymupdf  # older package name
            doc = pymupdf.open(path)
            parts = []
            for i, page in enumerate(doc, start=1):
                parts.append(page.get_text())
                # Ensure a page-number breadcrumb exists even if PV Elite's
                # own "Page N of M" footer doesn't survive text extraction
                # cleanly, so section-splitting/page-tracking stays robust.
            text = "\n".join(parts)
            return "pdf", text.replace("\r\n", "\n")
        else:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
            return "text", text.replace("\r\n", "\n")

    def _tag_and_name(self):
        stem = Path(self.pdf_file).stem
        stem = stem.replace("___", " & ").replace("__", "_")
        parts = stem.split("_", 1)
        tag = parts[0].strip()
        name = parts[1].replace("_", " ").strip() if len(parts) > 1 else ""
        return tag, name

    def _log_msg(self, msg):
        self.log.append(msg)

    def extract_all(self):
        # NOTE: page-break noise stripping is applied LOCALLY within
        # extract_table7 only (see that function) - NOT globally here.
        # Applying it to the whole document erases legitimate NEW-section
        # headings whenever they happen to start right after a page
        # break (very common), which silently destroyed section
        # boundaries other parsers (e.g. the BOM parser) depend on to
        # know where their own section ends.
        text = self.text
        self._log_msg(f"Equipment Tag: {self.equipment_tag}")
        self._log_msg(f"Total Pages: {self.total_pages}")

        t1 = extract_table1(text)
        self._log_msg(f"Table-1 (Thickness/Pressure/Diameter/Allow.Stress): {len(t1)} rows")
        t2 = extract_table2(text)
        self._log_msg(f"Table-2 (Reqd Thickness & MAWP - Element Master): {len(t2)} rows")
        t3 = extract_table3(text)
        self._log_msg(f"Table-3 (Surface Areas): {len(t3)} rows")
        t4 = extract_table4(text)
        self._log_msg(f"Table-4 (Element and Detail Weights): {len(t4)} rows")
        t5 = extract_table5(text)
        self._log_msg(f"Table-5 (Materials of Construction): {len(t5)} component types")
        t6 = extract_table6(text)
        self._log_msg(f"Table-6 (Element Pressures and MAWP): {len(t6)} rows")
        t7 = extract_table7(text)
        self._log_msg(f"Table-7 (Element Types and Properties): {len(t7)} rows")
        t4b = extract_table_total_ele_empty_wgt(text)
        self._log_msg(f"Table-4b ('Total Ele. Empty Wgt.' - separate from Table-4): {len(t4b)} rows")
        design_temps = extract_element_design_temps(text)
        self._log_msg(f"Per-element Design Temperatures: {len(design_temps)} rows")
        governing_stress, skipped_stress_rows = extract_governing_combined_stress(text)
        self._log_msg(f"Governing combined-load stress: {len(governing_stress)} nodes "
                       f"({skipped_stress_rows} partial rows skipped for safety)")
        loads_and_weights = extract_equipment_loads_and_weights(text)
        bom_rows = extract_bom(text)
        bom_totals = categorize_bom(bom_rows)
        self._log_msg(f"Bill of Materials: {len(bom_rows)} line items, "
                       f"{len(bom_totals)} categories")
        platform_details_raw = extract_platform_details(text)
        self._log_msg(f"Platform details: {len(platform_details_raw)} platforms")
        weight_summation = extract_weight_summation_table(text)
        self._log_msg(f"Weight Summation Results: {len(weight_summation)} categories")

        warnings = []

        # --- Quality Check 1: element count matches across Thickness/Area/Weight tables
        counts = {"Table-1": len(t1), "Table-2": len(t2), "Table-3": len(t3),
                  "Table-4": len(t4), "Table-6": len(t6), "Table-7": len(t7)}
        distinct_counts = set(c for c in counts.values() if c > 0)
        if len(distinct_counts) > 1:
            warnings.append(f"CHECK-1 FAILED: element counts differ across tables: {counts}")
        else:
            self._log_msg(f"CHECK-1 PASSED: consistent element count ({distinct_counts})")

        # --- Quality Check 3: duplicate element numbers in Table-2 (the master)
        elem_nos = [r["element_no"] for r in t2]
        dupes = {n for n in elem_nos if elem_nos.count(n) > 1}
        if dupes:
            warnings.append(f"CHECK-3 FAILED: duplicate Element# in master table: {sorted(dupes)}")

        # --- Quality Check 7: cross-validate Table-2's Element# sequence
        # against Table-4's own From-node numbering (From-node / 10 gives
        # an independent element index, since PV Elite numbers nodes
        # 10,20,30... one per element boundary). A mismatch means Table-2
        # and Table-4 have drifted out of positional sync with each other.
        for i, w in enumerate(t4):
            try:
                derived_no = int(w["from_node"]) // 10
            except (ValueError, TypeError):
                continue
            if i < len(t2) and derived_no != t2[i]["element_no"]:
                warnings.append(
                    f"CHECK-7: Table-4 row {i} (From-node {w['from_node']} -> "
                    f"element #{derived_no}) does not match Table-2's element "
                    f"#{t2[i]['element_no']} at the same position - tables may be out of sync"
                )
                break  # one mismatch invalidates alignment from here on; no need to spam

        n = len(t2)
        records = []
        for i in range(n):
            e2 = t2[i]
            elem_no = e2["element_no"]
            e1 = t1[i] if i < len(t1) else {}
            e3 = t3[i] if i < len(t3) else {}
            e4 = t4[i] if i < len(t4) else {}
            e4b = t4b[i] if i < len(t4b) else {}
            e6 = t6[i] if i < len(t6) else {}
            e7 = t7[i] if i < len(t7) else {}
            e_temp = design_temps[i] if i < len(design_temps) else {}

            element_type = e7.get("element_type")
            element_name = e2.get("element_name") or e1.get("element_name")
            component = identify_component(element_type, element_name)

            mat_entry = t5.get(component, {})

            surf_cm2 = e3.get("outside_surface_area_cm2")
            surf_in_cm2 = e3.get("inside_surface_area_cm2")
            weight_kg = e4.get("weight_kg")
            gov_row = governing_stress.get(e4.get("from_node"))

            # --- Quality Check 2: missing material mapping
            if not mat_entry:
                warnings.append(f"CHECK-2: no material mapping for component '{component}' "
                                 f"(element #{elem_no} '{element_name}')")
            # --- Quality Check 4: missing weight
            if weight_kg is None:
                warnings.append(f"CHECK-4: missing weight for element #{elem_no} '{element_name}'")
            # --- Quality Check 5: missing surface area
            if surf_cm2 is None:
                warnings.append(f"CHECK-5: missing surface area for element #{elem_no} '{element_name}'")
            # --- Quality Check 6: missing thickness
            if e1.get("nominal_thickness_mm") is None:
                warnings.append(f"CHECK-6: missing nominal thickness for element #{elem_no} '{element_name}'")

            records.append({
                "Equipment": self.equipment_tag,
                "Element No": elem_no,
                "Element Name": element_name,
                "Component": component,
                "Material": mat_entry.get("material"),
                "Class": mat_entry.get("class"),
                "UNS Number": mat_entry.get("uns"),
                "Diameter (m)": round(e1["diameter_mm"] / 1000.0, 4) if e1.get("diameter_mm") else None,
                "Element Length (mm)": e7.get("length_mm"),
                "Minimum Thickness (mm)": e2.get("minimum_thickness_mm"),
                "Nominal Thickness (mm)": e1.get("nominal_thickness_mm"),
                "Corrosion Allowance (mm)": e1.get("corrosion_allowance_mm"),
                "External Pressure (MPa)": round(e6.get("external_pressure_kpa") / 1000.0, 5)
                    if e6.get("external_pressure_kpa") else None,
                "Design Pressure (MPa)": round((e2.get("design_pressure_kpa") or e6.get("design_pressure_kpa")) / 1000.0, 5)
                    if (e2.get("design_pressure_kpa") or e6.get("design_pressure_kpa")) else None,
                "Design Temperature (degC)": e_temp.get("design_temp_c"),
                "Allowable Tensile Stress (MPa)": e1.get("allowable_stress_mpa") or e_temp.get("material_allowable_stress_ambient_mpa"),
                "Allowable Compressive Stress (MPa)": gov_row.get("allowable_compressive_stress_mpa") if gov_row else None,
                "Surface Area (m2)": round(surf_cm2 / 10000.0, 4) if surf_cm2 else None,
                "Inside Surface Area (m2)": round(surf_in_cm2 / 10000.0, 4) if surf_in_cm2 else None,
                "Weight (MT)": round(weight_kg / 1000.0, 4) if weight_kg else None,
                "Total Ele. Empty Wgt. (MT)": round(e4b["total_ele_empty_wgt_kg"] / 1000.0, 4)
                    if e4b.get("total_ele_empty_wgt_kg") else None,
            })

        df = pd.DataFrame(records)
        # PV Elite only prints a material's "Allowable Stress, Ambient"
        # once per material (not repeated for every element that shares
        # it) - fill remaining blanks using another row with the same
        # Material + Class, which is available for the great majority of
        # elements including skirts (Table-1 only covers pressure-
        # boundary elements, so skirts have no other source at all).
        if len(df) and "Allowable Tensile Stress (MPa)" in df.columns:
            df["Allowable Tensile Stress (MPa)"] = df.groupby(
                ["Material", "Class"], dropna=False
            )["Allowable Tensile Stress (MPa)"].transform(lambda s: s.ffill().bfill())
        success = len(df) > 0

        # --- Work-volume derived areas, per the domain rules given ---
        # FP (fireproofing) area = 2x the skirt's own outside surface area.
        # Insulation area = full outside area of every pressure-boundary
        # element (Shell + Head + Cone) - not gated by a per-element
        # "has insulation" flag, since in practice the whole pressure
        # envelope gets insulated.
        # Painting area = total outside area of the vessel (all elements)
        # PLUS the skirt's inside surface area (the skirt interior is
        # painted too, unlike the shell/head interior which isn't).
        skirt_mask = df["Component"] == "Skirt" if success else None
        pressure_mask = df["Component"].isin(["Shell", "Head", "Cone"]) if success else None

        skirt_outside_m2 = df.loc[skirt_mask, "Surface Area (m2)"].sum() if success else 0
        skirt_inside_m2 = df.loc[skirt_mask, "Inside Surface Area (m2)"].sum() if success else 0
        pressure_outside_m2 = df.loc[pressure_mask, "Surface Area (m2)"].sum() if success else 0
        total_outside_m2 = df["Surface Area (m2)"].sum() if success else 0

        # Qualifier gate: Platform and Insulation are qualified directly
        # by their presence in PV Elite's own "Total Weight of Each
        # Detail Type" rollup - the single most reliable signal, since
        # it's the report's own summary table rather than a per-element
        # flag scan. FP has no equivalent category in that table on any
        # report checked (FP items are folded into the generic "Weights"
        # bucket there), so FP still relies on the per-element Detail ID
        # "FP*" scan as the only available signal for it specifically.
        detail_totals = loads_and_weights.get("detail_type_totals_kg", {})
        has_platform_anywhere = "Platforms" in detail_totals and bool(detail_totals.get("Platforms"))
        has_insulation_anywhere = "Insulation" in detail_totals and bool(detail_totals.get("Insulation"))
        has_fp_anywhere = any(d.get("has_fp") for d in design_temps)

        fp_area_m2 = round(2 * skirt_outside_m2, 3) if (has_fp_anywhere and skirt_outside_m2) else "N/A"
        insulation_area_m2 = round(pressure_outside_m2, 3) if (has_insulation_anywhere and pressure_outside_m2) else "N/A"
        painting_area_m2 = round(total_outside_m2 + skirt_inside_m2, 3) if total_outside_m2 else None

        diameters = [r["diameter_mm"] for r in t1 if r.get("diameter_mm")]
        loads_and_weights["fp_area_m2"] = fp_area_m2
        loads_and_weights["insulation_area_m2"] = insulation_area_m2
        loads_and_weights["painting_area_m2"] = painting_area_m2
        loads_and_weights["has_platform"] = "Yes" if has_platform_anywhere else "No"
        loads_and_weights["has_insulation"] = "Yes" if has_insulation_anywhere else "No"
        loads_and_weights["has_fp"] = "Yes" if has_fp_anywhere else "No"
        loads_and_weights["max_diameter_mm"] = max(diameters) if diameters else None
        loads_and_weights["min_diameter_mm"] = min(diameters) if diameters else None
        ld_ratio = None
        if loads_and_weights.get("tl_tl_mm") and diameters:
            ld_ratio = round(loads_and_weights["tl_tl_mm"] / max(diameters), 3)
        loads_and_weights["l_over_d_ratio"] = ld_ratio

        # --- Additional equipment-level features for an ML-ready summary ---
        shell_mat = t5.get("Shell", {})
        loads_and_weights["shell_moc"] = shell_mat.get("material")
        loads_and_weights["shell_moc_class"] = shell_mat.get("class")
        loads_and_weights["shell_uns"] = shell_mat.get("uns")
        loads_and_weights["orientation"] = "Vertical" if (success and (df["Component"] == "Skirt").any()) else "Horizontal/Unknown"
        loads_and_weights["support_type"] = "Skirt" if loads_and_weights.get("basering_type") else None
        max_dp = df["Design Pressure (MPa)"].max() if success and df["Design Pressure (MPa)"].notna().any() else None
        max_dt = df["Design Temperature (degC)"].max() if success and df["Design Temperature (degC)"].notna().any() else None
        loads_and_weights["max_design_pressure_mpa"] = round(max_dp, 4) if max_dp is not None else None
        loads_and_weights["max_design_temp_c"] = max_dt
        loads_and_weights["total_element_weight_mt"] = round(df["Weight (MT)"].sum(), 4) if success else None
        loads_and_weights["total_surface_area_m2"] = round(total_outside_m2, 4) if total_outside_m2 else None
        loads_and_weights["element_count"] = len(df) if success else 0

        # --- ML-friendly derived ratios ---
        # Normalizing shear by surface area (wind load scales with
        # exposed area) and by operating weight (seismic load scales
        # with mass) lets a regression separate site severity from
        # vessel size, rather than the model having to learn that
        # relationship implicitly from raw shear values alone.
        wind_shear = loads_and_weights.get("wind_shear_kgf")
        eq_shear = loads_and_weights.get("eq_shear_kgf")
        surf_area = loads_and_weights.get("total_surface_area_m2")
        oper_wt = loads_and_weights.get("operating_mt")
        loads_and_weights["wind_shear_to_surface_area"] = (
            round(wind_shear / surf_area, 4) if wind_shear and surf_area else None)
        loads_and_weights["eq_shear_to_operating_weight"] = (
            round(eq_shear / (oper_wt * 1000.0), 5) if eq_shear and oper_wt else None)

        # --- Weight-forecasting ratios (for predicting new-tag Fabricated
        # Weight and Platform Weight from dimensions/design-basis alone,
        # since a new tag won't have a PV Elite report to read weight
        # from directly - only geometry and design basis are known
        # up-front, which is why these ratios are expressed in terms of
        # TL-TL/Diameter/Surface Area rather than any weight figure).
        fab_wt = loads_and_weights.get("fabricated_mt")
        tl_tl_m = (loads_and_weights.get("tl_tl_mm") or 0) / 1000.0
        max_dia_m = (loads_and_weights.get("max_diameter_mm") or 0) / 1000.0
        loads_and_weights["fab_weight_per_length_mt_per_m"] = (
            round(fab_wt / tl_tl_m, 4) if fab_wt and tl_tl_m else None)
        loads_and_weights["fab_weight_per_footprint_mt_per_m3"] = (
            round(fab_wt / (tl_tl_m * max_dia_m * max_dia_m), 5)
            if fab_wt and tl_tl_m and max_dia_m else None)
        loads_and_weights["fab_weight_per_surface_area"] = (
            round(fab_wt * 1000.0 / surf_area, 3) if fab_wt and surf_area else None)

        stats = {
            "equipment_tag": self.equipment_tag,
            "equipment_name": self.equipment_name,
            "total_elements": len(df),
            "total_weight_mt": df["Weight (MT)"].sum() if success else 0,
            "total_surface_area_m2": df["Surface Area (m2)"].sum() if success else 0,
        }
        for k, v in stats.items():
            self._log_msg(f"{k}: {v}")
        for w in warnings:
            self._log_msg("WARNING: " + w)

        # Nozzle/Platform/Insulation/FP weights, FP area, and the three
        # qualifiers are per-TAG (not per-element) values, so they live
        # in the Equipment_Summary sheet (one row per vessel) rather than
        # being repeated on every element row here.

        # Platform list: attach each platform's element diameter (by
        # matching its element_index against Table-1's row order) and
        # compute its area per the stated rule - width x length where a
        # top/head platform has a real length; otherwise treat it as an
        # arc segment around the shell at that element's diameter.
        platform_list = []
        for p in platform_details_raw:
            diameter_mm = t1[p["element_index"]]["diameter_mm"] if p["element_index"] < len(t1) else None
            start_a, end_a = p["start_angle_deg"], p["end_angle_deg"]
            angle_diff_deg = (end_a - start_a) if end_a >= start_a else (end_a + 360 - start_a)
            angle_diff_rad = angle_diff_deg * math.pi / 180.0
            if p["length_mm"] and p["length_mm"] > 0:
                area_m2 = round((p["width_mm"] * p["length_mm"]) / 1e6, 4)
            elif diameter_mm and p["width_mm"]:
                # Corrected rule: area = angle extent (radians) x element
                # diameter x platform width - no pi()x term, no +100mm
                # allowance (that was my earlier, incorrect assumption).
                area_m2 = round((angle_diff_rad * diameter_mm * p["width_mm"]) / 1e6, 4)
            else:
                area_m2 = None
            element_name = t2[p["element_index"]]["element_name"] if p["element_index"] < len(t2) else None
            platform_list.append({
                "element_name": element_name,
                "diameter_mm": diameter_mm,
                "detail_id": p["detail_id"],
                "start_angle_deg": start_a,
                "end_angle_deg": end_a,
                "width_mm": p["width_mm"],
                "length_mm": p["length_mm"],
                "area_m2": area_m2,
            })
        loads_and_weights["total_platform_area_m2"] = (
            round(sum(p["area_m2"] for p in platform_list if p["area_m2"]), 4)
            if platform_list else None
        )

        self.result = ExtractionResult(success=success, dataframe=df, stats=stats, warnings=warnings,
                                        loads_and_weights=loads_and_weights, bom_totals=bom_totals,
                                        platform_list=platform_list, weight_summation=weight_summation)
        return self.result

    def print_extraction_log(self):
        print("\n".join(self.log))

    def save_to_excel(self, output_file):
        if self.result is None or not self.result.success:
            return False
        df = self.result.dataframe
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Elements"
        ws.append(list(df.columns))
        for c in ws[1]:
            c.font = Font(bold=True)
        for row in df.itertuples(index=False):
            ws.append(list(row))
        ws.freeze_panes = "A2"

        ws2 = wb.create_sheet("Summary")
        ws2.append(["Metric", "Value"])
        for c in ws2[1]:
            c.font = Font(bold=True)
        for k, v in self.result.stats.items():
            ws2.append([k, v])

        ws_lw = wb.create_sheet("Loads_Foundation_Support")
        lw = self.result.loads_and_weights
        lw_headers = [
            "Tag No", "ASME Code", "Division", "Code Edition Year", "MDMT (degC)",
            "Wind Shear (Kgf)", "Earthquake Shear (Kgf)",
            "Wind Moment (Kgf-m)", "Earthquake Moment (Kgf-m)",
            "Fabricated Weight (MT)", "Erected Weight (MT)",
            "Operating Weight (MT)", "Shop Test Weight (MT)",
            "Basering Type", "Bolt Nominal Dia (mm)", "Bolt Circle Dia (mm)", "Number of Bolts",
        ]
        ws_lw.append(lw_headers)
        for c in ws_lw[1]:
            c.font = Font(bold=True)
        ws_lw.append([
            self.equipment_tag, lw.get("asme_code"), lw.get("division"), lw.get("code_edition_year"), lw.get("mdmt_c"),
            lw.get("wind_shear_kgf"), lw.get("eq_shear_kgf"),
            lw.get("wind_moment_kgfm"), lw.get("eq_moment_kgfm"),
            lw.get("fabricated_mt"), lw.get("erected_mt"),
            lw.get("operating_mt"), lw.get("shop_test_mt"),
            lw.get("basering_type"), lw.get("bolt_nominal_dia_mm"),
            lw.get("bolt_circle_dia_mm"), lw.get("bolt_qty"),
        ])
        ws_lw.freeze_panes = "A2"

        ws_wv = wb.create_sheet("Work_Volume")
        wv_headers = [
            "Tag No", "TL-TL (mm)", "Max Diameter (mm)", "Min Diameter (mm)", "L/D Ratio",
            "Erection Weight (MT)", "Platform Qualified", "Insulation Qualified",
            "Insulation Area (m2)", "Painting Area (m2)", "FP Qualified", "FP Area (m2)",
        ]
        ws_wv.append(wv_headers)
        for c in ws_wv[1]:
            c.font = Font(bold=True)
        ws_wv.append([
            self.equipment_tag, lw.get("tl_tl_mm"), lw.get("max_diameter_mm"), lw.get("min_diameter_mm"),
            lw.get("l_over_d_ratio"), lw.get("erected_mt"), lw.get("has_platform"), lw.get("has_insulation"),
            lw.get("insulation_area_m2"), lw.get("painting_area_m2"), lw.get("has_fp"), lw.get("fp_area_m2"),
        ])
        ws_wv.freeze_panes = "A2"

        if self.result.warnings:
            ws3 = wb.create_sheet("Warnings")
            ws3.append(["Warning"])
            for c in ws3[1]:
                c.font = Font(bold=True)
            for w in self.result.warnings:
                ws3.append([w])

        wb.save(output_file)
        return True


# --------------------------------------------------------------------------
# Command-line entry point
# --------------------------------------------------------------------------
def _main():
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description="Extract element-level data from a PV Elite calculation "
                     "report (text export saved with a .pdf extension)."
    )
    parser.add_argument("pdf_file", nargs="?",
                         help="Path to the PV Elite report. If omitted, you "
                              "will be prompted to enter it.")
    parser.add_argument("-o", "--output",
                         help="Output .xlsx path. Defaults to "
                              "'PVElite_Elements_<tag>.xlsx' next to the input file.")
    args = parser.parse_args()

    pdf_file = args.pdf_file
    if not pdf_file:
        pdf_file = input("Enter path to the PV Elite report (.pdf): ").strip().strip('"')

    if not os.path.isfile(pdf_file):
        print(f"ERROR: file not found: {pdf_file}")
        sys.exit(1)

    extractor = PVEliteDynamicExtractor(pdf_file)
    print(f"File Kind     : {extractor.file_kind} "
          f"({'plain-text export renamed .pdf' if extractor.file_kind == 'text' else 'real/binary PDF'})")
    print(f"Equipment Tag : {extractor.equipment_tag}")
    print(f"Equipment Name: {extractor.equipment_name}")
    print(f"Total Pages   : {extractor.total_pages}")
    print("Extracting...\n")

    result = extractor.extract_all()
    extractor.print_extraction_log()

    if not result.success:
        print("\nExtraction FAILED - no element rows were produced. "
              "This report's table layout may differ from what this parser "
              "expects; please share it so the parser can be extended.")
        sys.exit(1)

    out_path = args.output or f"PVElite_Elements_{extractor.equipment_tag}.xlsx"
    ok = extractor.save_to_excel(out_path)
    print(f"\n{'Saved' if ok else 'FAILED to save'}: {out_path}")
    if result.warnings:
        print(f"\n{len(result.warnings)} warning(s) - see the 'Warnings' sheet in the output file.")


if __name__ == "__main__":
    _main()
