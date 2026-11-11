import streamlit as st
import io
import html
import math
import re
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pydicom
import matplotlib.pyplot as plt
import numpy as np

def add_inline_bar(paragraph, img_buf, width_cm=3.0, text_pt=11):
    """Insert the percentile bar inline and drop it so the bar sits on the text's mid-line.
    A picture normally rests on the text baseline, which makes the bar float above the words."""
    import struct
    img_buf.seek(0)
    w_px, h_px = struct.unpack('>II', img_buf.read(24)[16:24])  # PNG header: width, height
    img_buf.seek(0)
    run = paragraph.add_run()
    run.add_picture(img_buf, width=Cm(width_cm))
    height_pt = width_cm * 28.346 * h_px / w_px
    lower_pt = max(0.0, height_pt / 2 - 0.35 * text_pt)   # bar centre -> centre of the text line
    pos = OxmlElement('w:position')
    pos.set(qn('w:val'), str(-int(round(lower_pt * 2))))  # value is in half-points
    run._r.get_or_add_rPr().append(pos)
    return run

def generate_percentile_bar_inline(percentile):
    """Slim percentile bar for use INSIDE a line of text (e.g. the EFW line).
    The picture is only as tall as the bar itself and the bar sits at its exact centre, so the
    bar lines up with the middle of the text instead of floating above the line."""
    if not percentile:
        return None
    try:
        pct = float(percentile)
    except:
        return None

    fig = plt.figure(figsize=(3, 0.28))
    fig.patch.set_alpha(0)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.patch.set_alpha(0)
    for a, b, col in [(0, 5, '#D32F2F'), (5, 10, '#FFC107'), (10, 90, '#4CAF50'), (90, 95, '#FFC107'), (95, 100, '#D32F2F')]:
        ax.plot([a, b], [0, 0], color=col, linewidth=6, zorder=1, solid_capstyle='butt')
    dot_color = '#000000' if 10 <= pct <= 90 else '#D32F2F'
    ax.scatter(min(max(pct, 0), 100), 0, color=dot_color, s=60, zorder=3, edgecolors='white', linewidths=1.5)
    for x in (10, 90):
        ax.plot([x, x], [-0.6, 0.6], color='black', linewidth=1, zorder=2)
    ax.set_xlim(-2, 102)
    ax.set_ylim(-1, 1)
    ax.axis('off')

    buf = io.BytesIO()
    plt.savefig(buf, format="png", transparent=True, dpi=250)  # fixed size: no cropping, bar stays centred
    plt.close(fig)
    buf.seek(0)
    return buf

def generate_percentile_bar(percentile):
    """Generates a color-coded PNG image of a percentile bar."""
    if not percentile:
        return None
    
    try:
        pct = float(percentile)
    except:
        return None

    # Make the figure slightly taller to fit the color zones better
    fig, ax = plt.subplots(figsize=(3, 0.5))
    fig.patch.set_alpha(0)
    
    # 1. Draw the color zones
    # Red: < 5th
    ax.plot([0, 5], [0, 0], color='#D32F2F', linewidth=6, zorder=1, solid_capstyle='butt')
    # Yellow: 5th - 10th
    ax.plot([5, 10], [0, 0], color='#FFC107', linewidth=6, zorder=1, solid_capstyle='butt')
    # Green: 10th - 90th
    ax.plot([10, 90], [0, 0], color='#4CAF50', linewidth=6, zorder=1, solid_capstyle='butt')
    # Yellow: 90th - 95th
    ax.plot([90, 95], [0, 0], color='#FFC107', linewidth=6, zorder=1, solid_capstyle='butt')
    # Red: > 95th
    ax.plot([95, 100], [0, 0], color='#D32F2F', linewidth=6, zorder=1, solid_capstyle='butt')
    
    # 2. Draw the marker dot
    # If the dot is in a red or yellow zone, make it stand out more
    dot_color = '#000000' if 10 <= pct <= 90 else '#D32F2F'
    ax.scatter(pct, 0, color=dot_color, s=60, zorder=3, edgecolors='white', linewidths=1.5)
    
    # 3. Add subtle tick marks at 10 and 90
    ax.plot([10, 10], [-0.15, 0.15], color='black', linewidth=1, zorder=2)
    ax.plot([90, 90], [-0.15, 0.15], color='black', linewidth=1, zorder=2)
    
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.3, 0.3)
    ax.axis('off')
    
    buf = io.BytesIO()
    # Increased DPI slightly for a crisper look in Word
    plt.savefig(buf, format="png", bbox_inches='tight', transparent=True, dpi=250)
    plt.close(fig)
    buf.seek(0)
    return buf

# ============================================================
# BARCELONA / FMF Z-SCORE CALCULATOR
# ============================================================

def _ga_days(ga_weeks: float) -> float:
    return ga_weeks * 7.0

def _z_to_percentile(z: float) -> float:
    sign = 1 if z >= 0 else -1
    x = abs(z) / math.sqrt(2.0)
    t = 1.0 / (1.0 + 0.3275911 * x)
    a1, a2, a3, a4, a5 = 0.254829592, -0.284496736, 1.421413741, -1.453152027, 1.061405429
    erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * math.exp(-x * x)
    cdf = 0.5 * (1.0 + sign * erf)
    return cdf * 100.0

def _pct_result(z: float):
    """z-score -> percentile (kept inside 0.1-99.9 so the report never shows 0.0 or 100.0)."""
    pct = min(max(_z_to_percentile(z), 0.1), 99.9)
    return {"z_score": round(z, 2), "percentile": round(pct, 1)}

def _ua_ref(ga_weeks: float):
    """(mean, SD) of umbilical artery PI at this GA - matched to the Barcelona calculator."""
    return 3.587744 - 0.137978 * ga_weeks + 0.0017778 * ga_weeks ** 2, 0.299

def _mca_ref(ga_weeks: float):
    """(mean, SD) of MCA PI at this GA - matched to the Barcelona calculator."""
    return (-0.005827 * ga_weeks ** 2 + 0.335068 * ga_weeks - 2.754777,
            -0.001350 * ga_weeks ** 2 + 0.082878 * ga_weeks - 0.893367)

def calculate_ua_pi(pi: float, ga_weeks: float):
    """Umbilical artery PI percentile, matched to the Barcelona Doppler calculator.
    Fitted to 8 Barcelona readings (20/26/32/38 wk, two PIs each) and confirmed at 3 off-grid checks:
    mean PI is quadratic in GA (weeks), SD constant ~0.30. Valid ~18-42 weeks."""
    if ga_weeks < 18 or ga_weeks > 42:
        return None
    mean, sd = _ua_ref(ga_weeks)
    return _pct_result((pi - mean) / sd)

def calculate_mca_pi(pi: float, ga_weeks: float):
    """MCA PI percentile, matched to the Barcelona Doppler calculator.
    Fitted to 17 Barcelona readings (max difference 0.4 percentile points) and confirmed at
    4 off-grid checks: mean PI and SD are both quadratic in GA (weeks). Fitted 20-41 weeks; valid ~19-42."""
    if ga_weeks < 19 or ga_weeks > 42:
        return None
    mean, sd = _mca_ref(ga_weeks)
    return _pct_result((pi - mean) / sd)

# CPR percentile is matched to the Barcelona calculator (set False to flag it with a dagger again).
CPR_PERCENTILE_VALIDATED = True

def _cpr_ref(ga_weeks: float):
    """(mean, SD) of the cerebro-placental ratio (MCA PI / UA PI) at this GA - matched to the Barcelona calculator.
    Barcelona models CPR as a normal distribution on the raw ratio; mean and SD are both quadratic in GA (weeks)."""
    mean = -4.0489 + 0.381875 * ga_weeks - 0.00588 * ga_weeks ** 2
    sd = -0.98616 + 0.091304 * ga_weeks - 0.00141 * ga_weeks ** 2
    return mean, sd

def calculate_cpr(cpr: float, ga_weeks: float):
    """Cerebro-placental ratio (MCA PI / UA PI) percentile, matched to the Barcelona Doppler calculator.
    Fitted to 34 Barcelona readings (20-40 weeks; worst difference 0.7 percentile points, i.e. rounding).
    Fitted 20-40 weeks; 41-42 weeks is extrapolated (unchecked). Valid ~20-42."""
    if cpr <= 0 or ga_weeks < 20 or ga_weeks > 42:
        return None
    mean, sd = _cpr_ref(ga_weeks)
    return _pct_result((cpr - mean) / sd)

PSV_MOM_ABNORMAL = 1.5      # MCA PSV at or above 1.5 MoM = risk of fetal anaemia

def mca_psv_median(ga_weeks: float):
    """Median MCA peak systolic velocity (cm/s) at this GA. Close to the Mari formula but fitted to the
    Barcelona calculator's readings (19 readings, MoM within ~0.01-0.02 of Barcelona's)."""
    return math.exp(2.3055 + 0.04625 * ga_weeks)

def calculate_mca_psv_mom(psv: float, ga_weeks: float):
    """MCA PSV (cm/s) -> multiples of the median (MoM). Barcelona rejects GA below ~19 weeks; fitted 20-40."""
    if psv <= 0 or ga_weeks < 20 or ga_weeks > 42:
        return None
    med = mca_psv_median(ga_weeks)
    return {"median": round(med, 1), "mom": round(psv / med, 2)}

def calculate_dv_pi(pi: float, ga_weeks: float):
    """Ductus venosus PI percentile, matched to the Barcelona Doppler calculator.
    Fitted to 24 Barcelona readings (20/24/28/32/36/40 wk, four PIs each; worst difference 0.6 percentile points,
    i.e. rounding) and confirmed on 10 off-grid readings (22-38 wk; worst 0.7): normal distribution on the raw PI,
    mean linear in GA (weeks), SD constant 0.149. Barcelona refuses 18, 19 and 42 weeks, so this is valid from
    20 weeks up to (not including) 41 weeks - 41 weeks onward was not checked."""
    if pi <= 0 or ga_weeks < 20 or ga_weeks >= 41:
        return None
    mean = 0.902388 - 0.011594 * ga_weeks
    return _pct_result((pi - mean) / 0.149302)

def calculate_uta_pi(pi: float, ga_weeks: float):
    """Mean uterine artery PI (average of right and left) percentile, fitted to the Barcelona calculator.
    Fitted to 20 Barcelona readings: log-normal, median and spread both quadratic in GA (weeks).
    17 of 20 readings within 1 percentile point; worst 3.5 points near the median at 34-36 weeks.
    Validated 20-41 weeks; below 20 weeks it is extrapolated (unchecked). Valid ~16-42 weeks."""
    if pi <= 0 or ga_weeks < 16 or ga_weeks > 42:
        return None
    ln_median = 8.68354e-4 * ga_weeks ** 2 - 7.75178e-2 * ga_weeks + 1.29545
    sdlog = 3.50871e-5 * ga_weeks ** 2 - 3.69546e-3 * ga_weeks + 0.297911
    return _pct_result((math.log(pi) - ln_median) / sdlog)

# ============================================================
# DICOM SR PARSER
# ============================================================
# --- BIOMETRY / ANOMALY-SCAN MEASUREMENTS ---
# SR concept name (lower-case) -> (key in latest_sr.json, keep one decimal in mm?)
# The key feeds <key> (size, mm), <key>_wks / <key>_dys (GA) and <key>_pct (percentile).
SR_BIOMETRY_KEYS = {
    'biparietal diameter': ('bpd', False), 'head circumference': ('hc', False),
    'abdominal circumference': ('ac', False), 'femur length': ('fl', False),
    'humerus length': ('hl', False),
    'radius length': ('radius', False), 'ulna length': ('ulna', False),
    'tibia length': ('tibia', False), 'fibula length': ('fibula', False),
    'trans cerebellar diameter': ('tcd', False), 'transcerebellar diameter': ('tcd', False),
    'anterior horn lateral ventricular width': ('lv', True), 'lateral ventricle': ('lv', True),
    'cisterna magna length': ('cm', True), 'cisterna magna': ('cm', True),
    'nuchal fold thickness': ('nf', True), 'nuchal fold': ('nf', True),
}
# sizes that are NOT handled by the original centimetre block
SR_EXTRA_SIZE_KEYS = {'radius', 'ulna', 'tibia', 'fibula', 'tcd', 'lv', 'cm', 'nf'}

def _sr_norm(concept):
    return re.sub(r'\s+', ' ', str(concept)).strip().lower()

def _sr_size_mm(num_val, unit, one_decimal):
    """SR length (cm or mm) -> millimetres as text. Small structures keep one decimal (5.2), the rest are whole mm."""
    if unit == 'centimeter': mm = num_val * 10
    elif unit == 'millimeter': mm = num_val
    else: return None
    return f"{mm:.1f}".rstrip('0').rstrip('.') if one_decimal else f"{int(mm)}"

def extract_sr_data(ds):
    from datetime import datetime, timedelta
    extracted = {}
    
    # We will use a flat list to grab every concept in the tree first, 
    # then parse them in order.
    flat_items = []
    
    VESSELS = ["Umbilical Artery", "Uterine Artery", "Fetal Doppler", "Maternal Doppler",
               "MCA", "Middle Cerebral Artery"]

    def walk_tree(seq, parent_site="", parent_side=""):
        for item in seq:
            concept = ""
            if 'ConceptNameCodeSequence' in item:
                concept = item.ConceptNameCodeSequence[0].get('CodeMeaning', '')
            
            new_site = parent_site
            new_side = parent_side
            # Catch any site identifier used by the machine
            if concept in VESSELS:
                new_site = concept
                new_side = ""
                # Laterality (Left / Right / Unilateral) is a CODE child of the vessel container
                for sub in item.get('ContentSequence', []):
                    if sub.get('ValueType', '') == 'CODE' and 'ConceptNameCodeSequence' in sub:
                        if sub.ConceptNameCodeSequence[0].get('CodeMeaning', '') == 'Laterality':
                            new_side = sub.ConceptCodeSequence[0].get('CodeMeaning', '').strip().lower()
            
            # Store the concept, the value sequence, the parent site and the side
            flat_items.append((concept, item, new_site, new_side))
            
            if 'ContentSequence' in item:
                walk_tree(item.ContentSequence, new_site, new_side)

    if 'ContentSequence' in ds:
        walk_tree(ds.ContentSequence)

    # Now we parse the flat list
    current_param = None
    for concept, item, site, side in flat_items:
        
        # Keep track of which biometry parameter we are under
        _bk = SR_BIOMETRY_KEYS.get(_sr_norm(concept))
        if _bk:
            current_param = _bk[0]

        # --- TEXT VALUES ---
        if item.ValueType == 'TEXT':
            val = item.get('TextValue', '')
            if val: 
                if concept == 'EDD from average ultrasound age': extracted['edd_cua'] = val
                if concept == 'EDD from LMP': extracted['edd_lmp'] = val
                if concept == 'LMP': extracted['lmp'] = val

        # --- NUMERIC VALUES ---
        if item.ValueType == 'NUM' and 'MeasuredValueSequence' in item:
            val_seq = item.MeasuredValueSequence[0]
            num_val = float(val_seq.get('NumericValue', 0))
            unit = val_seq.get('MeasurementUnitsCodeSequence', [{}])[0].get('CodeMeaning', '')
            
            # Biometry Sizes (in cm, convert to mm)
            if unit == 'centimeter':
                if concept == 'Biparietal Diameter': extracted['bpd'] = f"{int(num_val * 10)}"
                elif concept == 'Head Circumference': extracted['hc'] = f"{int(num_val * 10)}"
                elif concept == 'Abdominal Circumference': extracted['ac'] = f"{int(num_val * 10)}"
                elif concept == 'Femur Length': extracted['fl'] = f"{int(num_val * 10)}"
                elif concept == 'Humerus length': extracted['hl'] = f"{int(num_val * 10)}"
                elif concept == 'Gestational Sac Diameter': extracted['msd'] = f"{int(num_val * 10)}"
                elif concept == 'Crown Rump Length': extracted['crl'] = f"{int(num_val * 10)}"
                elif concept == 'Nuchal Translucency': extracted['nt'] = f"{int(num_val * 10)}"

            # Anomaly-scan sizes: TCD, lateral ventricle, cisterna magna, nuchal fold, radius, ulna, tibia, fibula
            if _bk and _bk[0] in SR_EXTRA_SIZE_KEYS:
                _mm = _sr_size_mm(num_val, unit, _bk[1])
                if _mm is not None:
                    extracted[_bk[0]] = _mm

            if concept == 'Nasal Bone':
                extracted['nasal_bone'] = "Present" if num_val > 0 else "Absent"
            
            # Doppler (Umbilical / Uterine / MCA): PI, RI and S/D, with Left/Right for uterine
            if concept in ('Pulsatility Index', 'Resistivity Index', 'Systolic to Diastolic Velocity Ratio'):
                metric = {'Pulsatility Index': 'pi', 'Resistivity Index': 'ri',
                          'Systolic to Diastolic Velocity Ratio': 'sd'}[concept]
                if "Umbilical" in site or "Fetal" in site: base = "umbilical"
                elif "Uterine" in site or "Maternal" in site: base = "uterine"
                elif "MCA" in site or "Middle Cerebral" in site: base = "mca"
                else: base = None
                if base:
                    val = str(round(num_val, 2))
                    if base == "uterine" and side in ("left", "right"):
                        extracted[f"uterine_{side}_{metric}"] = val
                    else:
                        extracted[f"{base}_{metric}"] = val

            # MCA peak systolic velocity (stored in cm/s)
            if concept == 'Peak Systolic Velocity' and ("MCA" in site or "Middle Cerebral" in site):
                _u = unit.lower()
                _v = num_val * 100 if ("meter per second" in _u and "centi" not in _u) else num_val
                extracted['mca_psv'] = str(round(abs(_v), 1))   # machines can store it negative

            # Summary & Ratios
            if concept == 'Estimated Weight': 
                extracted['fetal_weight'] = str(int(num_val))
            if concept == 'EFW percentile rank':
                extracted['fetal_weight_pct'] = str(round(num_val, 1))
            
            if concept == 'HC/AC': extracted['hc_ac'] = str(round(num_val, 2))
            if concept == 'FL/AC': extracted['fl_ac'] = str(round(num_val, 1))
            if concept == 'FL/BPD': extracted['fl_bpd'] = str(round(num_val, 1))
            if concept == 'FL/HC': extracted['fl_hc'] = str(round(num_val, 1))

            if concept == 'Fetal Heart Rate': 
                extracted['fhr'] = str(int(num_val))
            
            if concept == 'Composite Ultrasound Age': 
                extracted['ga_usg_days'] = int(num_val)
                extracted['mean_age'] = f"{int(num_val)//7}weeks {int(num_val)%7}days ± 2w"
            
            if concept == 'Gestational Age by LMP':
                extracted['ga_lmp_days'] = int(num_val)

            # Biometry Percentiles & GA
            if concept == 'Growth Percentile Rank':
                if current_param:
                    extracted[f"{current_param}_pct"] = str(round(num_val, 1))

            if concept == 'Gestational Age':
                if current_param:
                    extracted[f"{current_param}_wks"] = str(int(num_val) // 7)
                    extracted[f"{current_param}_dys"] = str(int(num_val) % 7)

        # Reset current_param when we hit a new group
        if concept in ['Biometry Group', 'Fetal Long Bones']:
            current_param = None
    
    # Uterine mean of left/right (unsided key = mean, same as the listener)
    for metric in ("pi", "ri", "sd"):
        vals = [float(extracted[f"uterine_{s}_{metric}"]) for s in ("left", "right")
                if f"uterine_{s}_{metric}" in extracted]
        if vals:
            extracted[f"uterine_{metric}"] = f"{sum(vals)/len(vals):.2f}"

    # Date calculations
    today = datetime.now()
    if 'lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['lmp'] = (today - timedelta(days=extracted['ga_lmp_days'])).strftime("%d-%m-%Y")
    if 'edd_lmp' not in extracted and 'ga_lmp_days' in extracted:
        extracted['edd_lmp'] = (today + timedelta(days=(280 - extracted['ga_lmp_days']))).strftime("%d-%m-%Y")
    if 'edd_cua' not in extracted and 'ga_usg_days' in extracted:
        extracted['edd_cua'] = (today + timedelta(days=(280 - extracted['ga_usg_days']))).strftime("%d-%m-%Y")

    # --- GRAB AND SPLIT PATIENT NAME/AGE ---
    if hasattr(ds, 'PatientName'):
        raw_name = str(ds.PatientName).replace('^', ' ').strip()
        extracted['patient_name'] = raw_name
        # Try to split the name and age (e.g., "SHIVANI 22Y/F")
        if 'Y/F' in raw_name.upper():
            parts = raw_name.upper().split('Y/F')
            age_part = parts[0].strip().split(' ')[-1] # Get the last word before Y/F (the age)
            name_part = parts[0].strip().replace(age_part, '').strip() # Remove the age from the name
            extracted['patient_age'] = age_part
            extracted['patient_name'] = name_part
    # ---------------------------------------

    return extracted
# ============================================================
# DOCX HELPERS
# ============================================================
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def set_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)

def add_run(paragraph, text, bold=False, italic=False, size=11, underline=False, color=None):
    run = paragraph.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    
    if color is not None:
        run.font.color.rgb = color
        # Force the XML to use the color
        rPr = run._element.get_or_add_rPr()
        c = OxmlElement('w:color')
        c.set(qn('w:val'), str(color))
        rPr.append(c)
        
    return run

def set_table_width(table, widths):
    """Fixed layout with an explicit total width (twips) so Word and LibreOffice draw the same width."""
    table.autofit = False
    total = sum(int(round(int(w) / 635)) for w in widths)   # EMU -> twips
    tblW = table._tbl.tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        table._tbl.tblPr.append(tblW)
    tblW.set(qn('w:w'), str(total))
    tblW.set(qn('w:type'), 'dxa')
    for i, w in enumerate(widths):
        table.columns[i].width = Emu(int(w))

_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]

def normalize_table_props(doc):
    """Put every table's <w:tblPr> children in the order the Word schema requires (some renderers ignore the rest)."""
    for tbl in doc.tables:
        pr = tbl._tbl.tblPr
        kids = list(pr)
        kids.sort(key=lambda e: _TBLPR_ORDER.index(e.tag.split('}')[1]) if e.tag.split('}')[1] in _TBLPR_ORDER else 99)
        for k in kids:
            pr.remove(k)
        for k in kids:
            pr.append(k)

def finish_document(doc):
    """Word needs a paragraph after a table that ends the document, and it makes one of normal height itself. When the
    last table ends near the page bottom that paragraph spills onto a blank page that cannot be deleted. Add our own,
    1 pt high, so it always fits."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    pPr = p._p.get_or_add_pPr()
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '2')
    rPr.append(sz)
    pPr.append(rPr)

def set_cell_background(cell, hex_color):
    """Sets the background shading of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

# Table colour themes (header, row 1, row 2): light shades to conserve ink. One per report, rotating.
COLOR_THEMES = [
    ("Green",  "A5D6A7", "E8F5E9", "D3EAD5"),
    ("Blue",   "BBDEFB", "E3F2FD", "D0E7FA"),
    ("Pink",   "F8BBD0", "FCE4EC", "F9D5E3"),
    ("Lavender", "CDBDEB", "F7F3FF", "ECE5F8"),
]
DISABLED_THEMES = {"Green"}      # kept in COLOR_THEMES but skipped by the rotation; remove the name here to switch it back on
ROTATION = [i for i, t in enumerate(COLOR_THEMES) if t[0] not in DISABLED_THEMES]
ACTIVE_THEME = ROTATION[0]
COLOR_DOPPLER = "COLOR_DOPPLER"                                   # dedicated Color Doppler template id
PINK_THEME_INDEX = [t[0] for t in COLOR_THEMES].index("Pink")      # Color Doppler reports are always pink
COLOR_STATE_FILE = "report_color_state.json"   # sits next to app.py; remembers which colour is next

def _color_state_path():
    import os
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), COLOR_STATE_FILE)

def load_color_index():
    """Colour theme to use for the current report (survives app restarts)."""
    import json
    try:
        with open(_color_state_path()) as f:
            idx = int(json.load(f).get("index", 0)) % len(COLOR_THEMES)
    except Exception:
        idx = ROTATION[0]
    return idx if idx in ROTATION else _next_enabled(idx)

def _next_enabled(idx):
    """Next colour after idx that is not disabled."""
    for step in range(1, len(COLOR_THEMES) + 1):
        j = (idx + step) % len(COLOR_THEMES)
        if j in ROTATION:
            return j
    return ROTATION[0]

def advance_color_index():
    """Called when a report is downloaded: the next report gets the next colour."""
    import json
    try:
        next_index = _next_enabled(load_color_index())   # read BEFORE the file is reopened for writing
        with open(_color_state_path(), "w") as f:
            json.dump({"index": next_index}, f)
    except Exception:
        pass

def get_gradient_color(row_index, total_rows=6):
    """Alternating table shading from the active colour theme."""
    _, header, row1, row2 = COLOR_THEMES[ACTIVE_THEME % len(COLOR_THEMES)]
    if row_index == 0:
        return header      # Header: darker shade
    elif row_index % 2 == 1:
        return row1        # Row 1: lighter
    else:
        return row2        # Row 2: slightly darker

def generate_afi_text(afi_val, is_total=True):
    try:
        val = float(afi_val)
    except (ValueError, TypeError):
        return "Adequate", False
    if not is_total: return "", False 
    if val == 0: return "Nil, Not a single pocket of even one cm seen", True
    elif 0 < val < 4: return "Moderate to severe oligohydramnios", True
    elif 4 <= val < 6: return "Mild to moderate oligohydramnios", True
    elif 6 <= val < 8: return "Mild oligohydramnios", True
    elif 8 <= val < 10: return "On the lower side", True
    elif 10 <= val <= 20: return "Adequate", False
    elif 20 < val < 25: return "On the higher side", True
    else: return "Polyhydramnios", True

# ============================================================
# DOCX BUILDER
# ============================================================
PLACENTA_OPTS = ["Anterior", "Posterior", "Fundo-Anterior", "Fundo-Posterior",
                 "Fundal extending antero-posteriorly"]

LIE_OPTIONS = ["None", "LOA", "ROA", "LOP", "ROP", "LOL", "ROL"]
LIE_LABELS = {
    "None": "None",
    "LOA": "Left Occiputo-Anterior", "ROA": "Right Occiputo-Anterior",
    "LOP": "Left Occiputo-Posterior", "ROP": "Right Occiputo-Posterior",
    "LOL": "Left Occiputo-Lateral", "ROL": "Right Occiputo-Lateral",
}

PRESENTATION_OPTIONS = ["CEPHALIC", "BREECH", "OBLIQUE LIE", "TRANSVERSE LIE", "VARIABLE"]
# Default presentation per template (can always be changed in the dropdown)
PRESENTATION_DEFAULTS = {"T3": "CEPHALIC", "T16": "VARIABLE", "COLOR_DOPPLER": "CEPHALIC", "ANOMALY": "VARIABLE"}

def _is_lie_presentation(data):
    return str(data.get('presentation', '') or '').upper() in ("OBLIQUE LIE", "TRANSVERSE LIE")

def presentation_text(data):
    """Wording for the top line and impression.
    Cephalic  -> 'Cephalic presentation & Left Occiputo-Anterior lie' (lie optional)
    Breech    -> 'Breech presentation'
    Oblique / Transverse -> 'Oblique Lie with head in Right upper quadrant'"""
    pres = str(data.get('presentation', 'variable') or 'variable').upper()
    if pres in ("OBLIQUE LIE", "TRANSVERSE LIE"):
        side = str(data.get('head_side', '') or '').strip()
        level = str(data.get('head_level', '') or '').strip().lower()
        quad = f" with head in {side} {level} quadrant" if side and level else ""
        return f"{pres.title()}{quad}"
    text = f"{pres.capitalize()} presentation"
    lie = data.get('lie', 'None')
    if pres == "CEPHALIC" and lie and lie != "None":
        text += f" & {LIE_LABELS.get(lie, lie)} lie"
    return text

def scar_text(val):
    """Scar thickness with grading: >2 mm Adequate, =2.0 mm Borderline adequate, <2 mm Reduced."""
    s = str(val if val is not None else "").strip()
    if not s:
        return ""
    m = re.fullmatch(r'(\d+(?:\.\d+)?)\s*(?:mm)?', s, re.I)
    if not m:
        return f"{s} mm"
    num = float(m.group(1))
    grade = "Adequate" if num > 2 else ("Reduced" if num < 2 else "Borderline adequate")
    return f"{m.group(1)} mm ({grade})"

def fmt_cm(val):
    """Cervical length in cm: '3.5' / '3.5cm' -> '3.5 cm'; '35 mm' -> '3.5 cm'.
    A bare number above 10 can only be millimetres, so it is converted. Text such as 'Normal' is left alone."""
    s = str(val if val is not None else "").strip()
    if not s:
        return ""
    m = re.fullmatch(r'(\d+(?:\.\d+)?)\s*(mm|cm)?', s, re.I)
    if not m:
        return s
    num = float(m.group(1))
    unit = (m.group(2) or "").lower()
    if unit == "mm" or (unit == "" and num > 10):
        num /= 10
    return f"{num:g} cm"

def _fmt_date(s):
    """04-01-2027 -> 04/01/2027"""
    s = str(s or "").strip()
    return re.sub(r'^(\d{1,2})-(\d{1,2})-(\d{4})$', r'\1/\2/\3', s)

def _split_ga(s):
    """'22w1d' / '19weeks 5days ± 2w' -> (22, 1) / (19, 5); None if not parseable."""
    m = re.search(r'(\d+)\s*w(?:eeks?)?\s*(?:(\d+)\s*d)?', str(s or '').lower())
    return (int(m.group(1)), int(m.group(2) or 0)) if m else None

SMALLER_FETUS_MILD = "Fetus is smaller for alleged GA"
SMALLER_FETUS_SEVERE = "Fetus is smaller for alleged GA - IUGR vs Wrong dates"

def smaller_fetus_remark(data):
    """GA by USG behind GA by LMP: 14-20 days -> mild remark, 21+ days -> IUGR vs wrong dates. Else ''."""
    lmp = str(data.get('lmp', '') or '').strip()
    if not lmp or lmp == "Not Sure":
        return ""
    l, u = _split_ga(data.get('ga_lmp')), _split_ga(data.get('mean_age'))
    if not l or not u:
        return ""
    diff = (l[0] * 7 + l[1]) - (u[0] * 7 + u[1])
    if diff >= 21:
        return SMALLER_FETUS_SEVERE
    if diff >= 14:
        return SMALLER_FETUS_MILD
    return ""

def dating_rows(data):
    """Rows for the Dating table: (label, LMP text, GA weeks text, GA days text, EDD)."""
    def ga_cells(ga):
        if not ga: return "-", "-"
        w, d = ga
        return f"{w} Weeks", f"{d} {'Day' if d == 1 else 'Days'}"
    rows = []
    lmp = str(data.get('lmp', '') or '').strip()
    if lmp and lmp != "Not Sure":
        wk, dy = ga_cells(_split_ga(data.get('ga_lmp')))
        rows.append(("By LMP", f"LMP: {_fmt_date(lmp)}", wk, dy, _fmt_date(data.get('edd_lmp', '')) or "-"))
    wk, dy = ga_cells(_split_ga(data.get('mean_age') or data.get('ga_crl')))
    rows.append(("By USG", "", wk, dy, _fmt_date(data.get('edd_cua', '')) or "-"))
    return rows

def add_dating_table(doc, data, full_width=False):
    rows = dating_rows(data)
    tbl = doc.add_table(rows=1 + len(rows), cols=5)
    set_table_borders(tbl)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER if full_width else WD_TABLE_ALIGNMENT.LEFT
    widths = [Cm(3.4)] * 5 if full_width else [Cm(2.6), Cm(4.0), Cm(2.4), Cm(2.4), Cm(3.0)]
    if full_width:
        set_table_width(tbl, widths)
    for r in tbl.rows:
        for i, w in enumerate(widths):
            r.cells[i].width = w
    for i, text in {0: "Dating", 1: "LMP", 4: "EDD"}.items():
        c = tbl.cell(0, i)
        set_cell_background(c, get_gradient_color(0))
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(c.paragraphs[0], text, bold=True, size=10)
    ga_head = tbl.cell(0, 2).merge(tbl.cell(0, 3))
    set_cell_background(ga_head, get_gradient_color(0))
    ga_head.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(ga_head.paragraphs[0], "GA", bold=True, size=10)
    for ri, vals in enumerate(rows, start=1):
        for ci, t in enumerate(vals):
            c = tbl.cell(ri, ci)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(c.paragraphs[0], t, size=10)
    return tbl

def dating_preview(data):
    rows = dating_rows(data)
    out = [f"  {'Dating':<9}{'LMP':<20}{'GA':<20}{'EDD'}"]
    for label, lmp, wk, dy, edd in rows:
        out.append(f"  {label:<9}{lmp:<20}{(wk + ' ' + dy):<20}{edd}")
    return out

DOPPLER_KEYS = [f"{vsl}_{m}" for vsl in ("umbilical", "uterine", "uterine_left", "uterine_right", "mca")
                for m in ("pi", "ri", "sd")] + ["mca_psv"]

def generate_doppler_text(data):
    doppler_lines = []
    if data.get("umbilical_pi"):
        doppler_lines.append(f"Umbilical Artery: PI = {data.get('umbilical_pi')}")
    if data.get("uterine_pi"):
        notching = data.get("uterine_notching", "No notching")
        if notching == "No notching":
            doppler_lines.append(f"Uterine Artery: PI = {data.get('uterine_pi')} (No notching seen)")
        else:
            doppler_lines.append(f"Uterine Artery: PI = {data.get('uterine_pi')} ({notching})")
    if data.get("mca_pi"):
        doppler_lines.append(f"MCA: PI = {data.get('mca_pi')}")
    return "\n".join(doppler_lines)

def _pf(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def doppler_status(data):
    """Interpret each Doppler index from its percentile.
    Umbilical / uterine PI above the 95th = Raised; MCA PI / CPR below the 5th = Reduced; notching = abnormal.
    'all_normal' needs UA, MCA and uterine values all present and all within range."""
    def high(v):
        return None if v is None else ("Raised" if v > 95 else "Normal")
    def low(v):
        return None if v is None else ("Reduced" if v < 5 else "Normal")
    def psv(v):
        return None if v is None else ("Raised" if v >= PSV_MOM_ABNORMAL else "Normal")
    s = {
        "psv": psv(_pf(data.get("mca_psv_mom"))),
        "ua": high(_pf(data.get("umbilical_pct"))),
        "mca": low(_pf(data.get("mca_pct"))),
        "cpr": low(_pf(data.get("cpr_pct"))),
        "uta": high(_pf(data.get("uterine_pct"))),
        "uta_left": high(_pf(data.get("uterine_left_pct"))),
        "uta_right": high(_pf(data.get("uterine_right_pct"))),
        "notching": str(data.get("uterine_notching", "No notching")) != "No notching",
    }
    needed = (s["ua"], s["mca"], s["uta"])
    flags = (s["ua"], s["mca"], s["psv"], s["cpr"], s["uta"], s["uta_left"], s["uta_right"])
    s["all_normal"] = (all(x is not None for x in needed)
                       and all(x in (None, "Normal") for x in flags)
                       and not s["notching"])
    return s

NORMAL_DOPPLER_LINE = "Fetal & Maternal vessels reveal normal color flow characteristics."

def _add_color_doppler_section(doc, data):
    """Page 1 of the Color Doppler report (below the biometry): banner, Doppler table, notes."""
    from docx.enum.table import WD_ALIGN_VERTICAL
    status = doppler_status(data)
    doc.add_paragraph()

    W = [Cm(4.6), Cm(1.5), Cm(1.5), Cm(1.5), Cm(3.1), Cm(5.1)]   # 17.3 cm, full page width
    OK, BAD = RGBColor(0x2E, 0x7D, 0x32), RGBColor(0xC6, 0x28, 0x28)

    # --- banner ---
    band = doc.add_table(rows=1, cols=1)
    set_table_borders(band)
    band.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_width(band, [sum(W, Cm(0))])
    bc = band.cell(0, 0)
    bc.width = sum(W, Cm(0))
    set_cell_background(bc, get_gradient_color(0))
    bp = bc.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(bp, "COLOR DOPPLER STUDY", bold=True, size=13)
    bp2 = bc.add_paragraph()
    bp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(bp2, "Fetal & Maternal Vessels", italic=True, size=10)
    doc.add_paragraph()

    # --- table ---
    tbl = doc.add_table(rows=1, cols=6)
    set_table_borders(tbl)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_width(tbl, W)         # grid widths: Word and LibreOffice both honour these
    for i, text in enumerate(["Vessel", "S/D", "RI", "PI", "Percentile", "Interpretation"]):
        c = tbl.cell(0, i)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        c.width = W[i]
        set_cell_background(c, get_gradient_color(0))
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(c.paragraphs[0], text, bold=True, size=10)

    state = {"n": 1}

    def pct_cell(cell, pct, dagger=False):
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if not pct:
            add_run(cell.paragraphs[0], "-", size=10)
            return
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        add_run(p, f"{pct}%" + ("\u2020" if dagger else ""), bold=True, size=9)
        bar = generate_percentile_bar(pct)
        if bar:
            p2 = cell.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_after = Pt(0)
            p2.add_run().add_picture(bar, width=Cm(2.4))

    def interp_cell(cell, text):
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if not text:
            add_run(cell.paragraphs[0], "-", size=10)
        else:
            add_run(cell.paragraphs[0], text, bold=(text != "Normal"), size=10, color=(OK if text == "Normal" else BAD))

    def add_row(label, sd, ri, pi, pct, interp, dagger=False):
        cells = tbl.add_row().cells
        bg = get_gradient_color(state["n"])
        for i, c in enumerate(cells):
            c.width = W[i]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_background(c, bg)
        add_run(cells[0].paragraphs[0], label, bold=True, size=10)
        for i, t in ((1, sd), (2, ri), (3, pi)):
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cells[i].paragraphs[0], t if t else "-", size=10)
        pct_cell(cells[4], pct, dagger)
        interp_cell(cells[5], interp)
        state["n"] += 1

    def has(*keys):
        return any(data.get(k) for k in keys)

    def g(prefix, m):
        return data.get(f"{prefix}_{m}")

    if has("umbilical_pi", "umbilical_ri", "umbilical_sd"):
        add_row("Umbilical Artery", g("umbilical", "sd"), g("umbilical", "ri"), g("umbilical", "pi"),
                data.get("umbilical_pct"), status["ua"])
    if has("mca_pi", "mca_ri", "mca_sd"):
        add_row("Middle Cerebral Artery", g("mca", "sd"), g("mca", "ri"), g("mca", "pi"),
                data.get("mca_pct"), status["mca"])

    has_l = has("uterine_left_pi", "uterine_left_ri", "uterine_left_sd")
    has_r = has("uterine_right_pi", "uterine_right_ri", "uterine_right_sd")
    if has_l and has_r:
        add_row("Rt. Uterine Artery", g("uterine_right", "sd"), g("uterine_right", "ri"), g("uterine_right", "pi"),
                data.get("uterine_right_pct"), status["uta_right"])
        add_row("Lt. Uterine Artery", g("uterine_left", "sd"), g("uterine_left", "ri"), g("uterine_left", "pi"),
                data.get("uterine_left_pct"), status["uta_left"])
        add_row("Mean Uterine Artery", g("uterine", "sd"), g("uterine", "ri"), g("uterine", "pi"),
                data.get("uterine_pct"), status["uta"])
    elif has_l or has_r:
        side, label = ("uterine_left", "Lt. Uterine Artery") if has_l else ("uterine_right", "Rt. Uterine Artery")
        add_row(label, g(side, "sd"), g(side, "ri"), g(side, "pi"), data.get("uterine_pct"), status["uta"])
    elif has("uterine_pi", "uterine_ri", "uterine_sd"):
        add_row("Uterine Artery", g("uterine", "sd"), g("uterine", "ri"), g("uterine", "pi"),
                data.get("uterine_pct"), status["uta"])

    # --- MCA PSV row: velocity spans the S/D, RI and PI columns, MoM sits in the percentile column ---
    if data.get("mca_psv"):
        cells = tbl.add_row().cells
        bg = get_gradient_color(state["n"])
        for i, c in enumerate(cells):
            c.width = W[i]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_background(c, bg)
        add_run(cells[0].paragraphs[0], "MCA PSV", bold=True, size=10)
        lp = cells[0].add_paragraph()
        lp.paragraph_format.space_after = Pt(0)
        add_run(lp, "(Peak Systolic Velocity)", italic=True, size=8)
        span = cells[1].merge(cells[3])
        span.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(span.paragraphs[0], f"{data.get('mca_psv')} cm/s", bold=True, size=11)
        pc = cells[4]
        pc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if data.get("mca_psv_mom"):
            pc.paragraphs[0].paragraph_format.space_after = Pt(0)
            add_run(pc.paragraphs[0], f"{data.get('mca_psv_mom')} MoM", bold=True, size=11)
            if data.get("mca_psv_median"):
                p2 = pc.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.space_after = Pt(0)
                add_run(p2, f"(median {data.get('mca_psv_median')} cm/s)", italic=True, size=8)
        else:
            add_run(pc.paragraphs[0], "-", size=10)
        interp_cell(cells[5], status["psv"])
        state["n"] += 1

    # --- CPR row: ratio spans the S/D, RI and PI columns ---
    if data.get("cpr"):
        cells = tbl.add_row().cells
        bg = get_gradient_color(state["n"])
        for i, c in enumerate(cells):
            c.width = W[i]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_background(c, bg)
        add_run(cells[0].paragraphs[0], "Cerebro-Placental Ratio", bold=True, size=10)
        span = cells[1].merge(cells[3])
        span.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(span.paragraphs[0], f"{data.get('cpr')}", bold=True, size=11)
        sp2 = span.add_paragraph()
        sp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp2.paragraph_format.space_after = Pt(0)
        add_run(sp2, "(MCA PI \u00f7 UA PI)", italic=True, size=8)
        pct_cell(cells[4], data.get("cpr_pct"), dagger=not CPR_PERCENTILE_VALIDATED)
        interp_cell(cells[5], status["cpr"])
        state["n"] += 1

    # --- notes ---
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    add_run(p, "Uterine artery notching: ", bold=True)
    add_run(p, str(data.get("uterine_notching", "No notching")), italic=True)

    p = doc.add_paragraph()
    add_run(p, "Interpretation key: ", bold=True, size=8)
    add_run(p, "umbilical and uterine artery PI above the 95th percentile, and MCA PI or cerebro-placental ratio "
               "below the 5th percentile, are considered abnormal."
               + (" MCA PSV of 1.5 MoM or more is also considered abnormal (risk of fetal anaemia)."
                  if data.get("mca_psv") else ""), italic=True, size=8)
    if data.get("cpr_pct") and not CPR_PERCENTILE_VALIDATED:
        p = doc.add_paragraph()
        add_run(p, "\u2020 CPR percentile is provisional.", italic=True, size=8)

def _add_page_two_break(doc):
    """'P.T.O' marker, then start page 2 with the same top spacer as page 1."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(10)
    add_run(p, "P.T.O", bold=True, italic=True)
    for i, pt in enumerate((10, 10, 5, 5)):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.line_spacing = 1.0
        if i == 0:
            sp.paragraph_format.page_break_before = True
        r = sp.add_run("\u200B")
        r.font.name = "Calibri"
        r.font.size = Pt(pt)

def build_docx_bytes(data, template_id):
    global ACTIVE_THEME
    if template_id == ANOMALY:            # Anomaly Scan has its own two-page blue layout
        return build_anomaly_docx_bytes(data)
    ACTIVE_THEME = (PINK_THEME_INDEX if template_id == COLOR_DOPPLER
                    else int(data.get("color_theme", ROTATION[0])) % len(COLOR_THEMES))
    doc = Document()
    # T3 with the scar-thickness line: the extra line would push the report onto a second page, so the blank lines
    # between sections shrink to 6 pt. The top spacer, the bottom margin and every other template are untouched.
    _tight = template_id in ("T3", "T16")
    # Rare worst case (scar-thickness line AND a Doppler table): claw back a few more points so it still fits
    _tighter = _tight and ((bool(data.get("measure_scar")) and any(data.get(k) for k in DOPPLER_KEYS))
                           or (template_id == "T3" and bool(doppler_advice(data))))   # the advice paragraph needs the room
    def _gap():
        p = doc.add_paragraph()
        if _tight:
            p.paragraph_format.line_spacing = Pt(3)
        return p
    def _blank():
        """One full blank line (10 pt) above a heading."""
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = Pt(12)
        r = p.add_run("\u200B")
        r.font.size = Pt(10)
        return p
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    # === FLEXIBLE TOP SPACER ===
    
    # 2 blank lines (10pt) to push the report down
    for _ in range(1 if _tight else 2):
        _p = doc.add_paragraph()
        _p.paragraph_format.space_after = Pt(0)
        _p.paragraph_format.space_before = Pt(0)
        _p.paragraph_format.line_spacing = 1.0
        # Adding a zero-width space so Word doesn't merge the empty paragraphs
        _run = _p.add_run("\u200B") 
        _run.font.name = "Calibri"
        _run.font.size = Pt(10)
        
    # 2 tiny blank lines (5pt) for fine-tuning
    for _ in range(1 if _tighter else 2):
        _p = doc.add_paragraph()
        _p.paragraph_format.space_after = Pt(0)
        _p.paragraph_format.space_before = Pt(0)
        _p.paragraph_format.line_spacing = 1.0
        _run = _p.add_run("\u200B")
        _run.font.name = "Calibri"
        _run.font.size = Pt(5)
        
    # === END SPACER ===

    # 1. HEADER TABLE
    demo_table = doc.add_table(rows=2, cols=2)
    demo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(demo_table)
    for row in demo_table.rows:
        row.cells[0].width = Cm(8.5)
        row.cells[1].width = Cm(8.5)
    add_run(demo_table.cell(0, 0).paragraphs[0], f"NAME :- {data.get('patient_name', '')}", bold=True)
    add_run(demo_table.cell(0, 1).paragraphs[0], f"DATE :- {data.get('report_date', '')}", bold=True)
    add_run(demo_table.cell(1, 0).paragraphs[0], f"AGE/SEX :- {data.get('patient_age', '')}Y/ {data.get('p_sex', 'F')}", bold=True)
    add_run(demo_table.cell(1, 1).paragraphs[0], f"REF. BY:- {data.get('referred_by', '')}", bold=True)
    
    # 8pt gap
    p_gap1 = doc.add_paragraph()
    p_gap1.paragraph_format.space_after = Pt(0)
    p_gap1.paragraph_format.line_spacing = Pt(5 if _tight else 8)

    # 2. TITLE
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_p, "COLOR DOPPLER OBSTETRICS REPORT" if template_id == COLOR_DOPPLER else "ULTRASOUND OBSTETRICS REPORT",
            bold=True, underline=True, size=14, color=RGBColor(0x1F, 0x4E, 0x79))
    
    # 5pt gap
    p_gap2 = doc.add_paragraph()
    p_gap2.paragraph_format.space_after = Pt(0)
    p_gap2.paragraph_format.line_spacing = Pt(3 if _tight else 5)
    
    # 5pt gap
    p_gap3 = doc.add_paragraph()
    p_gap3.paragraph_format.space_after = Pt(0)
    p_gap3.paragraph_format.line_spacing = Pt(3 if _tight else 5)

    # 3. EARLY OBS
    if template_id == "EARLY_OBS":
        early_table = doc.add_table(rows=10, cols=2)
        set_table_borders(early_table)
        for row in early_table.rows:
            row.cells[0].width = Cm(5.5)
            row.cells[1].width = Cm(11.5)
        subchorionic = "Absent" if not data.get("subchorionic_present") else f"Present; {data.get('subchorionic_grade', 'Grade-I')}"
        corpus = "Not Appreciated" if not data.get("corpus_present") else f"Appreciated; {data.get('corpus_size', '')}mm in {data.get('corpus_side', 'Right')} Ovary"
        cardiac = f"Present, {data.get('fhr_val', '')} bpm" if data.get('cardiac_status') == "Present" else data.get('cardiac_status')
        
        early_data = [
            ("G-Sac", data.get('gsac', 'Seen')), ("CRL", data.get('crl', '')),
            ("Yolk sac", data.get('yolk', 'Not Seen')), ("Cardiac pulsations", cardiac),
            ("Decidual reaction", "Present"), ("Subchorionic collection", subchorionic),
            ("Corpus leuteal cyst", corpus), ("Cervix", fmt_cm(data.get('cervix', 'Normal'))),
            ("Internal Os", "Closed"), ("Placenta", "-"),
        ]
        for i, (label, val) in enumerate(early_data):
            add_run(early_table.cell(i, 0).paragraphs[0], label, italic=True, size=10)
            add_run(early_table.cell(i, 1).paragraphs[0], val, italic=True, size=10)
        doc.add_paragraph()

        add_dating_table(doc, data)

    # 3.5 NT_NB SCAN
    elif template_id == "NT_NB":
        nt_table = doc.add_table(rows=4, cols=4)
        set_table_borders(nt_table)
        for row in nt_table.rows:
            row.cells[0].width = Cm(4.0); row.cells[1].width = Cm(4.0); row.cells[2].width = Cm(4.0); row.cells[3].width = Cm(4.0)
            
        nt_data = [
            ("CRL", f"{data.get('crl', '')}mm", "NT", f"{data.get('nt', '')}mm"),
            ("Nasal Bone", data.get('nasal_bone', ''), "FHR", f"{data.get('fhr_val', '')} bpm"),
            ("DV PI", data.get('dv_pi', ''), "DV A-Wave", data.get('dv_awave', '')),
            ("Placenta", data.get('placenta_pos', ''), "Cervix", fmt_cm(data.get('cervix', ''))),
        ]
        
        for i, (l1, v1, l2, v2) in enumerate(nt_data):
            add_run(nt_table.cell(i, 0).paragraphs[0], l1, bold=True, size=10)
            add_run(nt_table.cell(i, 1).paragraphs[0], v1, italic=True, size=10)
            add_run(nt_table.cell(i, 2).paragraphs[0], l2, bold=True, size=10)
            add_run(nt_table.cell(i, 3).paragraphs[0], v2, italic=True, size=10)
            
        doc.add_paragraph()
        
        add_dating_table(doc, data)

    # 4. STANDARD OB
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Combine Presentation and Lie
        pres_text = presentation_text(data)
            
        add_run(p, f"Uterus shows single live fetus. Fetus is in ")
        add_run(p, f"{pres_text}", bold=True, italic=True)
        add_run(p, f" at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        
        # 8pt gap
        p_gap4 = doc.add_paragraph()
        p_gap4.paragraph_format.space_after = Pt(0)
        p_gap4.paragraph_format.line_spacing = Pt(4 if _tight else 8)

        bio_head = doc.add_paragraph()
        add_run(bio_head, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True)
        bio_head.paragraph_format.space_after = Pt(2 if _tight else 6)

        # --- NEW COMPACT COLORED BIOMETRY TABLE ---
        bio_table = doc.add_table(rows=6, cols=5)
        set_table_borders(bio_table)
        bio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        if template_id == COLOR_DOPPLER:
            set_table_width(bio_table, [Cm(3.4)] * 5)     # full page width, same as the Doppler tables
        
        # 1. Header Row
        headers = ["Fetal growth parameters", "mm", "Weeks", "Days", "Percentile"]
        widths = [Cm(3.4)] * 5 if template_id == COLOR_DOPPLER else [Cm(4.5), Cm(1.5), Cm(1.5), Cm(1.5), Cm(5.0)]  # Color Doppler: full width
        for i, text in enumerate(headers):
            cell = bio_table.cell(0, i)
            cell.width = widths[i]
            set_cell_background(cell, get_gradient_color(0))
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
            add_run(p, text, bold=True, size=10)

        # 2. Data Rows
        params = ["bpd", "hc", "ac", "fl", "hl"]
        param_names = ["Biparietal Diameter", "Head circumference", "Abdominal Circumference", "Femoral Length", "Humerus Length"]
        
        for row_idx, param in enumerate(params):
            row_i = row_idx + 1 # Offset by 1 for header
            size = data.get(f"{param}_size", ""); wks = data.get(f"{param}_wks", ""); dys = data.get(f"{param}_dys", "")
            
            # Row background color
            bg_color = get_gradient_color(row_i)
            
            # Col 0: Parameter Name
            cell0 = bio_table.cell(row_i, 0)
            cell0.width = widths[0]
            set_cell_background(cell0, bg_color)
            add_run(cell0.paragraphs[0], param_names[row_idx], size=10)
            
            # Col 1: Size (mm)
            cell1 = bio_table.cell(row_i, 1)
            cell1.width = widths[1]
            set_cell_background(cell1, bg_color)
            cell1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cell1.paragraphs[0], size if size else "", size=10)
            
            # Col 2: Weeks
            cell2 = bio_table.cell(row_i, 2)
            cell2.width = widths[2]
            set_cell_background(cell2, bg_color)
            cell2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cell2.paragraphs[0], wks if wks else "", size=10)
            
            # Col 3: Days
            cell3 = bio_table.cell(row_i, 3)
            cell3.width = widths[3]
            set_cell_background(cell3, bg_color)
            cell3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(cell3.paragraphs[0], dys if dys else "", size=10)
            
            # Col 4: Percentile Chart (Text next to the bar to save space)
            cell4 = bio_table.cell(row_i, 4)
            cell4.width = widths[4]
            set_cell_background(cell4, bg_color)
            cell4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            pct_val = data.get(f"{param}_pct", "")
            if pct_val and str(pct_val).strip() != "":
                try:
                    bar_img = generate_percentile_bar(pct_val)
                except:
                    bar_img = None
                if bar_img:
                    p = cell4.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.0
                    
                    # 1. Add the percentage text
                    add_run(p, f"{pct_val}% ", bold=True, size=9)
                    
                    # 2. Add the image (Reduced to 2.5cm to fit perfectly)
                    run = p.add_run()
                    run.add_picture(bar_img, width=Cm(2.5))
        # --- END BIOMETRY TABLE ---
                        
        # (Removed doc.add_paragraph() here to make them flush)

        # --- DYNAMIC RATIO TABLE BASED ON GESTATIONAL AGE ---
        import re
        mean_age_str = str(data.get('mean_age', '')).lower()
        
        # Find the digits immediately before "weeks"
        match = re.search(r'(\d+)\s*weeks', mean_age_str)
        
        if match:
            ga_weeks = int(match.group(1))
        else:
            ga_weeks = 0
        
        ratio_table = doc.add_table(rows=1, cols=2)
        ratio_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        ratio_table.columns[0].width = Cm(7.5); ratio_table.columns[1].width = Cm(8.5)
        if template_id == COLOR_DOPPLER:
            set_table_width(ratio_table, [Cm(7.86), Cm(7.86)])
            for _c, _w in zip(ratio_table.rows[0].cells, (Cm(7.86), Cm(7.86))):
                _c.width = _w
        
        if ga_weeks <= 22 and ga_weeks > 0:
            # Before 22 weeks
            add_run(ratio_table.cell(0, 0).paragraphs[0], "FL/HC= ", bold=True)
            add_run(ratio_table.cell(0, 0).paragraphs[0], f"{data.get('fl_hc', '')}", italic=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], "HC/AC = ", bold=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], f"{data.get('hc_ac', '')} (BOTH WNL)", italic=True)
        else:
            # 23 weeks and beyond
            add_run(ratio_table.cell(0, 0).paragraphs[0], "FL/AC= ", bold=True)
            add_run(ratio_table.cell(0, 0).paragraphs[0], f"{data.get('fl_ac', '')}", italic=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], "FL/BPD = ", bold=True)
            add_run(ratio_table.cell(0, 1).paragraphs[0], f"{data.get('fl_bpd', '')} (BOTH WNL)", italic=True)
        
        # (Removed doc.add_paragraph() here to make them flush)

        p = doc.add_paragraph()
        add_run(p, "MEAN AGE : ", bold=True); add_run(p, f"{data.get('mean_age', '')}", italic=True)

        # Color Doppler: page 1 = biometry + Doppler table (then P.T.O); page 2 = dating, weight, AFI ... impression
        dop_layout = template_id == COLOR_DOPPLER and any(data.get(k) for k in DOPPLER_KEYS)
        if dop_layout:
            _add_color_doppler_section(doc, data)
            _add_page_two_break(doc)
            add_dating_table(doc, data, full_width=True)
            doc.add_paragraph()

        # --- FETAL WEIGHT WITH BAR ---
        p = doc.add_paragraph()
        add_run(p, "FETAL WEIGHT : ", bold=True)
        
        # Calculate the ± 15% error margin
        fw_str = f"{data.get('fetal_weight', '')}g"
        try:
            fw_val = float(data.get('fetal_weight', ''))
            error_margin = int(round(fw_val * 0.15))
            fw_str = f"{int(fw_val)} grams ± {error_margin}gms"
        except:
            pass # If no weight is entered, just leave it blank
            
        add_run(p, fw_str, italic=True)
        if data.get("fetal_weight_pct"):
            p.add_run("  (Percentile: ")
            p.add_run(f"{data.get('fetal_weight_pct')}%)  ")
            bar_img_efw = generate_percentile_bar_inline(data.get("fetal_weight_pct"))
            if bar_img_efw:
                add_inline_bar(p, bar_img_efw, width_cm=3.0)

        p = doc.add_paragraph()
        add_run(p, "AFI:- ", bold=True)
        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        
        # We must define is_abn here so the bolding works!
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', '')
            desc, is_abn = generate_afi_text(afi_val, is_total=True)
            if desc == "Nil, Not a single pocket of even one cm seen":
                afi_str = desc
            else:
                afi_str = f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', '')
            try:
                afi_float = float(afi_val)
                if 2.0 <= afi_float <= 8.0:
                    afi_str = f"Adequate, largest pocket measures {afi_val} cm"
                    is_abn = False
                elif afi_float < 2.0:
                    afi_str = f"Reduced, largest pocket measures {afi_val} cm"
                    is_abn = True
                else:
                    afi_str = f"Increased, largest pocket measures {afi_val} cm"
                    is_abn = True
            except:
                afi_str = f"largest pocket measures {afi_val} cm"
                is_abn = False
                
        add_run(p, afi_str, bold=is_abn, italic=is_abn)
        p = doc.add_paragraph()
        add_run(p, "Foetal heart rate ", bold=True); add_run(p, f"is approx. {data.get('fhr', '')}.", italic=True)
        if not dop_layout:
            _gap()

            # --- DATING TABLE (LMP / USG: GA and EDD) ---
            add_dating_table(doc, data)
            _gap()

        # --- PLACENTA ---
        p = doc.add_paragraph()
        add_run(p, "Placenta ", bold=True)
        plac_pos = data.get('placenta_pos', 'Anterior')
        plac_os = data.get('placenta_os', 'away from the Os')
        
        # Add the bold italic position
        add_run(p, "is ", italic=False)
        add_run(p, plac_pos.lower(), bold=True, italic=True)
        
        if plac_os == "Placenta Praevia":
            plac_str = f" and Placenta Praevia ({data.get('praevia_grade', 'Grade I')})."
        elif plac_os == "Low-lying":
            dist = data.get('placenta_dist', '')
            plac_str = f" and Low-lying, about {dist}cm away from the Internal Os." if dist.strip() else f" and Low-lying, but away from the Internal Os."
        else:
            plac_str = f" and {plac_os}."
        add_run(p, plac_str)

        # --- CERVIX ---
        p = doc.add_paragraph()
        add_run(p, "CERVIX- ", bold=True)
        _cx = fmt_cm(data.get('cervix', ''))
        add_run(p, f"normal ({_cx}), internal Os is {data.get('internal_os', 'closed').lower()}." if _cx
                else f"normal, internal Os is {data.get('internal_os', 'closed').lower()}.")

        # --- CORD LOOP ---
        p = doc.add_paragraph()
        if data.get('cord_around') or data.get('cord_straddling'):
            if data.get('cord_around'): add_run(p, "Single loop of cord around the neck.", bold=True, italic=True)
            if data.get('cord_straddling'): add_run(p, "Single loop of cord straddling the neck.", bold=True, italic=True)
        else:
            add_run(p, "No loop of cord seen around the neck.", bold=True, italic=True)

        # --- SCAR THICKNESS (T3 only) ---
        if template_id in ("T3", COLOR_DOPPLER) and data.get('measure_scar'):
            p = doc.add_paragraph()
            add_run(p, "Scar Thickness: ", bold=True)
            add_run(p, scar_text(data.get('scar_size', '')), italic=True)

        # --- DOPPLER SECTION WITH DYNAMIC TABLE ---
        if template_id == COLOR_DOPPLER:
            pass    # drawn on page 1, under the biometry (see dop_layout above)
        elif any(data.get(k) for k in DOPPLER_KEYS):
            _ = _blank() if _tight else _gap()
            doppler_head = doc.add_paragraph()
            add_run(doppler_head, "DOPPLER INDICES:", bold=True, underline=True)
            
            # Table: Vessel | PI | RI | S/D | Notching | Percentile
            dop_table = doc.add_table(rows=1, cols=6)
            set_table_borders(dop_table)
            dop_table.alignment = WD_TABLE_ALIGNMENT.LEFT
            
            # 3.6 + 1.7 + 1.7 + 1.7 + 2.5 + 2.8 = 14.0 cm (matches Biometry table)
            dop_widths = [Cm(3.6), Cm(1.7), Cm(1.7), Cm(1.7), Cm(2.5), Cm(2.8)]
            
            # 1. Header Row
            headers = ["Vessel", "PI", "RI", "S/D", "Notching", "Percentile"]
            for i, text in enumerate(headers):
                cell = dop_table.cell(0, i)
                cell.width = dop_widths[i]
                set_cell_background(cell, get_gradient_color(0))
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run(cell.paragraphs[0], text, bold=True, size=10)
            
            # 2. Build the rows for each vessel
            row_idx = 1
            def add_doppler_row(vessel_name, pi_val, ri_val, sd_val, notching_text, pct_val):
                nonlocal row_idx
                row_cells = dop_table.add_row().cells
                bg_color = get_gradient_color(row_idx)
                
                texts = [vessel_name, pi_val, ri_val, sd_val, notching_text]
                for i, t in enumerate(texts):
                    row_cells[i].width = dop_widths[i]
                    set_cell_background(row_cells[i], bg_color)
                    if i > 0:
                        row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_run(row_cells[i].paragraphs[0], t if t else "-", size=10)
                
                # Percentile (Text + Bar)
                row_cells[5].width = dop_widths[5]
                set_cell_background(row_cells[5], bg_color)
                row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if pct_val:
                    try:
                        bar_img = generate_percentile_bar(pct_val)
                        if bar_img:
                            p = row_cells[5].paragraphs[0]
                            p.paragraph_format.space_after = Pt(0)
                            add_run(p, f"{pct_val}% ", bold=True, size=9)
                            p.add_run().add_picture(bar_img, width=Cm(2.0))
                    except:
                        pass
                else:
                    add_run(row_cells[5].paragraphs[0], "-", size=10)
                
                row_idx += 1
            
            def _has(*keys):
                return any(data.get(k) for k in keys)
            
            # Add the specific vessels
            if _has("umbilical_pi", "umbilical_ri", "umbilical_sd"):
                add_doppler_row("Umbilical Artery", data.get("umbilical_pi"), data.get("umbilical_ri"),
                                data.get("umbilical_sd"), "", data.get("umbilical_pct"))
            
            has_left = _has("uterine_left_pi", "uterine_left_ri", "uterine_left_sd")
            has_right = _has("uterine_right_pi", "uterine_right_ri", "uterine_right_sd")
            notching = data.get("uterine_notching", "No notching")

            def _ut_row(label, prefix, notch, pct):
                add_doppler_row(label, data.get(f"{prefix}_pi"), data.get(f"{prefix}_ri"),
                                data.get(f"{prefix}_sd"), notch, pct)

            if has_left and has_right:
                _ut_row("Lt. Uterine Artery", "uterine_left", "", "")
                _ut_row("Rt. Uterine Artery", "uterine_right", "", "")
                if _has("uterine_pi", "uterine_ri", "uterine_sd"):
                    _ut_row("Uterine Artery (Mean)", "uterine", notching, data.get("uterine_pct"))
            elif has_left or has_right:
                # Only one side measured: show it by name, carrying the notching and percentile
                if has_left:
                    _ut_row("Lt. Uterine Artery", "uterine_left", notching, data.get("uterine_pct"))
                else:
                    _ut_row("Rt. Uterine Artery", "uterine_right", notching, data.get("uterine_pct"))
            elif _has("uterine_pi", "uterine_ri", "uterine_sd"):
                _ut_row("Uterine Artery", "uterine", notching, data.get("uterine_pct"))

            if _has("mca_pi", "mca_ri", "mca_sd"):
                add_doppler_row("MCA", data.get("mca_pi"), data.get("mca_ri"), data.get("mca_sd"),
                                "", data.get("mca_pct"))

    # 5. IMPRESSION
    _ = _blank() if _tight else _gap()
    imp_head = doc.add_paragraph()
    add_run(imp_head, "IMPRESSION:", bold=True, underline=True, color=RGBColor(0x1F, 0x4E, 0x79))
    mean_age = data.get('mean_age', '')
    if template_id == "EARLY_OBS":
        if data.get('cardiac_status') == "Blighted Ovum":
            imp_text = f"INTRA UTERINE EMPTY GESTATIONAL SAC OF {data.get('ga_crl', '')} +/-1w WITH NO APPRECIABLE FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY FAILURE (BLIGHTED OVUM)."
        elif data.get('cardiac_status') == "Not yet appreciable":
            imp_text = f"Intra uterine pregnancy of {data.get('ga_crl', '')} +/-1w with not yet appreciable cardiac activity. Adv- Review scan after 2 weeks."
        else:
            imp_text = f"Intra uterine live pregnancy of {data.get('ga_crl', '')} +/-1w with good cardiac activity."
    elif template_id == "NT_NB":
        imp_text = f"Single live fetus of approx. {mean_age} gestation age. NT measures {data.get('nt', '')}mm. Nasal bone is {data.get('nasal_bone', '').lower()}."
    else:
        pres_text = presentation_text(data)
            
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {pres_text}{',' if _is_lie_presentation(data) else ''} of approx. {mean_age} gestation age."
    p = doc.add_paragraph(); add_run(p, imp_text, bold=True, italic=True)
    _imp_p = p
    if template_id == COLOR_DOPPLER and doppler_status(data)["all_normal"]:
        _blank()
        p = doc.add_paragraph(); add_run(p, NORMAL_DOPPLER_LINE, bold=True, italic=True)
    _adv = doppler_advice(data) if template_id in ("T3", COLOR_DOPPLER) else ""
    _sf = smaller_fetus_remark(data) if template_id not in ("EARLY_OBS", "NT_NB") else ""
    if _sf:                                    # fetus smaller than LMP dates: remark after a blank line, bold italics
        _b = _blank()
        p = doc.add_paragraph(); add_run(p, _sf, bold=True, italic=True)
        if _adv:                               # keep the chain impression -> remark -> advice -> PNDT box together
            _b.paragraph_format.keep_with_next = True
            p.paragraph_format.keep_with_next = True
            _imp_p.paragraph_format.keep_with_next = True
    if _adv:                                   # pathological Doppler: advice after a blank line, bold italics
        _blank().paragraph_format.keep_with_next = True
        p = doc.add_paragraph(); add_run(p, _adv, bold=True, italic=True)
        p.paragraph_format.keep_with_next = True
        # If the page is too full for impression + advice + PNDT box, they move to the next page TOGETHER
        # (instead of the PNDT box being left alone on a second page).
        imp_head.paragraph_format.keep_with_next = True
        _imp_p.paragraph_format.keep_with_next = True

    # 6. PNDT & NOTE COMBINED BOX
    _gap().paragraph_format.keep_with_next = bool(_adv)
    disc_table = doc.add_table(rows=1, cols=1)
    disc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(disc_table)
    cell = disc_table.cell(0, 0)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # First Line: PNDT Disclaimer
    add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=(6.5 if _tighter else 7 if _tight else 8))
    
    # Add a line break, then the Note
    _br = p.add_run(); _br.add_break()
    if _tight:
        _br.font.size = Pt(7)          # the line-break run would otherwise count as a full 11 pt line
    add_run(p, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=(6.5 if _tighter else 7 if _tight else 8))

    finish_document(doc)
    normalize_table_props(doc)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ============================================================
# ANOMALY SCAN TEMPLATE  (blue theme, 2 pages, own builder)
# ============================================================
ANOMALY = "ANOMALY"
AN_NAVY, AN_LIGHT, AN_LINE, AN_MID = "1F4E79", "DDEBF7", "9DC3E6", "2E75B6"
AN_FONT = "Arial"
AN_SHOW_DV = True                      # Ductus Venosus row: percentile is calculated from the typed PI (override possible)
AN_CONTENT_W = 17.2                    # cm (A4 21 cm less 1.9 cm margins each side)

# ---------- text helpers ----------
def _an_weeks(v):
    s = str(v or "").strip()
    return f"{s} weeks" if s else ""

def _an_days(v):
    s = str(v or "").strip()
    if not s:
        return ""
    return f"{int(s):02d} days" if s.isdigit() else f"{s} days"

def _an_mm(v):
    s = str(v or "").strip()
    return s if (not s or s.lower().endswith("mm")) else f"{s}mm"

def _an_ga_text(ga):
    return f"{ga[0]} weeks {ga[1]:02d} days" if ga else "\u2014"

def _an_date(s):
    s = _fmt_date(s)
    return s if s else "\u2014"

def anomaly_placenta_text(data):
    """e.g. 'Posterior, away from the Internal Os.'  (typist picks position + relation to Os)"""
    pos = str(data.get("placenta_pos", "") or "").strip()
    rel = data.get("placenta_os", "away from the Os")
    if rel == "Low-lying":
        dist = str(data.get("placenta_dist", "") or "").strip()
        r = "low-lying" + (f", {dist} cm from the Internal Os" if dist else "")
    elif rel == "Placenta Praevia":
        r = ("Placenta praevia, " + str(data.get("praevia_grade", "") or "")).rstrip(", ")
    else:
        r = "away from the Internal Os"
    return f"{pos}, {r}." if pos else ""

def anomaly_liquor(data):
    """(text, is_abnormal). Normal liquor is pre-filled as 'adequate'; the typist only supplies the number."""
    if data.get("afi_method", "Largest Pocket") == "Total 4 Quadrants":
        val = str(data.get("afi_total", "") or "").strip()
        desc, abn = generate_afi_text(val, is_total=True) if val else ("Adequate", False)
        if desc.startswith("Nil"):
            return f"AFI:- {desc}", True
        return f"AFI:- {desc if abn else desc.lower()} (total of four quadrants {val or '____'}cm)", abn
    val = str(data.get("afi_pocket", "") or "").strip()
    desc, abn = "adequate", False
    try:
        f = float(val)
        if f < 2.0:
            desc, abn = "Reduced", True
        elif f > 8.0:
            desc, abn = "Increased", True
    except ValueError:
        pass
    return f"AFI:- {desc} (largest pocket {val or '____'}cm)", abn

def anomaly_efw(data):
    fw = str(data.get("fetal_weight", "") or "").strip()
    sd = str(data.get("efw_sd", "") or "").strip()
    if fw and not sd:
        try:
            sd = str(int(round(float(fw) * 0.15)))      # same +/-15% convention as the other templates
        except ValueError:
            pass
    return f"{fw or '____'} gms +/- {sd or '____'} gms"

def anomaly_presentation_row(data):
    return presentation_text(data).replace(" presentation", "")

def anomaly_cervix(data):
    cv = fmt_cm(data.get("cervix", "")) or "____ cm"
    os_state = str(data.get("internal_os", "Closed") or "Closed").lower()
    return f"Internal Os is {os_state}, Cervical length- normal ({cv})."

def anomaly_opinion(data):
    ga = _split_ga(data.get("mean_age"))
    age = f"{ga[0]} weeks {ga[1]:02d} days \u00b1 2w" if ga else str(data.get("mean_age", "") or "____")
    pres = presentation_text(data).lower()
    return f"Single live intra-uterine foetus in {pres}{',' if _is_lie_presentation(data) else ''} of Approx. {age} gestational age."

def anomaly_doppler_rows(data):
    """[(label, PI value, percentile, remark, remark_is_abnormal)] - UA, mean uterine, (MCA if present), DV."""
    s = doppler_status(data)
    notch = str(data.get("uterine_notching", "No notching"))
    rows = []

    def rem(flags, ok="(WNL)."):
        return (", ".join(flags), True) if flags else (ok, False)

    v = data.get("umbilical_pi", "")
    rows.append(("UMBILICAL ARTERY PI", v, data.get("umbilical_pct", ""),
                 *(rem(["Raised PI"] if s["ua"] == "Raised" else []) if v else ("", False))))
    v = data.get("uterine_pi", "")
    fl = (["Raised PI"] if s["uta"] == "Raised" else []) + ([notch] if notch != "No notching" else [])
    rows.append(("MEAN UTERINE ARTERY PI", v, data.get("uterine_pct", ""), *(rem(fl) if v else ("", False))))
    v = data.get("mca_pi", "")
    if v:
        rows.append(("MCA PI", v, data.get("mca_pct", ""), *rem(["Reduced PI"] if s["mca"] == "Reduced" else [])))
    if AN_SHOW_DV:
        v = data.get("dv_pi", "")
        pct = _pf(data.get("dv_pct"))
        aw = str(data.get("dv_awave", "Positive") or "Positive")
        fl = (["Raised PI"] if (pct is not None and pct > 95) else []) + ([f"A-wave {aw.lower()}"] if aw != "Positive" else [])
        rows.append(("DUCTUS VENOSUS PI", v, data.get("dv_pct", ""),
                     *(rem(fl, "(WNL, Normal S, D & A Waves)") if v else ("", False))))
    return rows


AN_ADVICE_UP_TO_34 = ("Adv- Prompt obstetric consultation is advised. Close fetal surveillance with serial Doppler assessment and "
                      "repeat assessment within 3\u20137 days. The above recommendation is applicable to the current gestational "
                      "period up to 34 weeks, subject to stable maternal and fetal status. Earlier reassessment and intervention "
                      "may be necessary in case of deterioration.")
AN_ADVICE_BEYOND_34 = ("Adv- Urgent obstetric consultation is advised. Close fetal surveillance with Doppler assessment and "
                       "reassessment within 24\u201372 hours as clinically indicated. Delivery planning is advised, preferably "
                       "within 36\u201337+6 weeks where indicated by the Doppler abnormality and overall clinical assessment. "
                       "Earlier intervention may be necessary in case of fetal or maternal compromise.")

def _advice_for_ga(data):
    """Up to 34 weeks (34w0d) -> prompt consultation; beyond 34 weeks -> urgent consultation + delivery planning."""
    ga = _split_ga(data.get("mean_age"))
    if ga:
        days = ga[0] * 7 + ga[1]
    else:
        days = (_pf(data.get("ga_weeks_dop")) or 0) * 7
    return AN_ADVICE_BEYOND_34 if days > 34 * 7 else AN_ADVICE_UP_TO_34

def anomaly_advice(data):
    """Advice line under the opinion - only when any Doppler index is pathological."""
    if not any(r[4] for r in anomaly_doppler_rows(data)):
        return ""
    return _advice_for_ga(data)

def doppler_advice(data):
    """T3 / Color Doppler impression: advice when the mean uterine artery or umbilical artery PI is above the 95th
    percentile, the MCA PI or CPR is below the 5th percentile, or the MCA PSV is 1.5 MoM or more."""
    s = doppler_status(data)
    if not any(s[k] in ("Raised", "Reduced") for k in ("ua", "uta", "mca", "cpr", "psv")):
        return ""
    return _advice_for_ga(data)

# (label, data key prefix, has percentile)  -- per table
AN_CRANIUM = [("B.P.D", "bpd", True), ("HEAD CIRCUMFERENCE", "hc", True), ("CEREBELLUM (TCD)", "tcd", True),
              ("LATERAL VENTRICLE", "lv", False), ("CISTERNA MAGNA", "cm", False), ("NUCHAL FOLD", "nf", False)]
AN_ABDOMEN = [("ABDOMINAL CIRCUMFERENCE", "ac", True)]
AN_UPPER = [("HUMERUS", "hl", True), ("ULNA", "ulna", True), ("RADIUS", "radius", False)]
AN_LOWER = [("FEMUR", "fl", True), ("TIBIA", "tibia", True), ("FIBULA", "fibula", False)]
AN_INDICES = [("FL / AC RATIO", "fl_ac", "(20% - 24%)"), ("FL / BPD RATIO", "fl_bpd", "(71% - 87%)"),
              ("FL / HC RATIO", "fl_hc", "(18.60% - 20.40%)"), ("HC / AC RATIO", "hc_ac", "(1.04 - 1.22)")]

AN_CRANIUM_TEXT = ["Mid line falx seen. Both lateral ventricles are normal.", "The posterior fossa appears normal.",
                   "The cerebellum is normal.", "No identifiable intracranial lesion seen."]
AN_SPINE_TEXT = ["The foetal spine appears normal with normal curvature.",
                 "Entire spine is visualised in longitudinal and transverse axis.",
                 "The vertebrae and spinal canal appear normal."]
AN_FACE_TEXT = ["Foetal face seen in coronal and profile views.", "Both orbits and nose appear normal."]
AN_THORAX_TEXT = ["The heart appears in mid-thoracic location.", "The four chamber view and 3 vessel views are normal.",
                  "Both lungs seen.", "No evidence of pleural or pericardial effusion.", "No evidence of SOL in the thorax."]
AN_ABDOMEN_TEXT = ["Stomach and bladder are normal.", "The portal vein is normal.",
                   "No evidence of ascites. No abdominal wall defect is appreciated."]
AN_PNDT = ('"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the '
           'purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure '
           'has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by '
           'words, signs, or in any other manner."')

# ---------- docx helpers ----------
def _an_run(p, text, bold=False, italic=False, size=9.5, color=None):
    r = p.add_run(text)
    r.font.name = AN_FONT
    r._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), AN_FONT)
    r.font.size = Pt(size)
    r.bold, r.italic = bold, italic
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r

def _an_flat(p, before=0, after=0, line=None):
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if line:
        pf.line_spacing = Pt(line)

def _an_spacer(doc, pt=5.5, page_break=False):
    """Blank line of the given font size (zero-width space keeps the paragraph alive)."""
    p = doc.add_paragraph()
    _an_flat(p, line=round(pt * 1.2, 1))
    if page_break:
        p.paragraph_format.page_break_before = True
    _an_run(p, "\u200B", size=pt)

def _an_gap(doc, line=6.5):
    p = doc.add_paragraph()
    _an_flat(p, line=line)

def _an_lines(doc, texts):
    for t in texts:
        p = doc.add_paragraph()
        _an_flat(p, after=1, line=12.6)
        _an_run(p, t)

def _an_heading(doc, text):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()                       # border first, so python-docx keeps the schema order
    bd = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, val in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", AN_MID)):
        b.set(qn(f"w:{k}"), val)
    bd.append(b)
    pPr.append(bd)
    _an_flat(p, before=7, after=3)
    p.paragraph_format.keep_with_next = True
    _an_run(p, text, bold=True, size=10.5, color=AN_NAVY)

def _an_table(doc, rows, widths_cm):
    t = doc.add_table(rows=rows, cols=len(widths_cm))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    pr = t._tbl.tblPr
    bd = OxmlElement("w:tblBorders")
    for edge, sz, col in (("top", 6, AN_NAVY), ("left", 6, AN_NAVY), ("bottom", 6, AN_NAVY), ("right", 6, AN_NAVY),
                          ("insideH", 4, AN_LINE), ("insideV", 4, AN_LINE)):
        e = OxmlElement(f"w:{edge}")
        for k, val in (("val", "single"), ("sz", str(sz)), ("space", "0"), ("color", col)):
            e.set(qn(f"w:{k}"), val)
        bd.append(e)
    pr.append(bd)
    mar = OxmlElement("w:tblCellMar")
    for side, w in (("top", 40), ("left", 100), ("bottom", 40), ("right", 100)):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(w))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    pr.append(mar)
    widths = [Cm(w) for w in widths_cm]
    set_table_width(t, widths)
    for r in t.rows:
        trPr = r._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        for i, w in enumerate(widths):
            r.cells[i].width = w
    return t

def _an_cell(cell, text="", bold=False, italic=False, color=None, fill=None, align=None, size=9):
    from docx.enum.table import WD_ALIGN_VERTICAL
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if fill:
        set_cell_background(cell, fill)
    p = cell.paragraphs[0]
    _an_flat(p)
    if align is not None:
        p.alignment = align
    if text:
        _an_run(p, text, bold=bold, italic=italic, size=size, color=color)
    return p

def _an_label(cell, text):
    _an_cell(cell, text, bold=True, color=AN_NAVY, fill=AN_LIGHT)

def _an_pct(cell, pct):
    """'50.0% [bar]' - same percentile bar image the Doppler / biometry tables already use."""
    p = _an_cell(cell, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    s = str(pct or "").strip()
    if not s:
        return
    bar = generate_percentile_bar_inline(s)      # slim bar, vertically centred on the text line
    if bar:
        _an_run(p, f"{s}%  ", bold=True, size=8.5)
        add_inline_bar(p, bar, width_cm=2.8, text_pt=8.5)

def _an_measure_table(doc, spec, data):
    C = WD_ALIGN_PARAGRAPH.CENTER
    t = _an_table(doc, len(spec), [5.6, 1.9, 2.1, 1.9, 5.7])
    for i, (label, k, has_pct) in enumerate(spec):
        c = t.rows[i].cells
        _an_label(c[0], label)
        _an_cell(c[1], _an_mm(data.get(f"{k}_size", "")), align=C)
        _an_cell(c[2], _an_weeks(data.get(f"{k}_wks", "")), align=C)
        _an_cell(c[3], _an_days(data.get(f"{k}_dys", "")), align=C)
        if has_pct:
            _an_pct(c[4], data.get(f"{k}_pct", ""))
        else:
            _an_cell(c[4])
    return t

def anomaly_edd_rows(data):
    lmp_ok = str(data.get("lmp", "") or "").strip() not in ("", "Not Sure")
    ega_lmp = _an_ga_text(_split_ga(data.get("ga_lmp"))) if lmp_ok else "\u2014"
    edd_lmp = _an_date(data.get("edd_lmp")) if lmp_ok else "\u2014"
    ega_usg = _an_ga_text(_split_ga(data.get("mean_age")))
    return [("EGA BY LMP", ega_lmp, "EDD (LMP)", edd_lmp), ("EGA BY USG", ega_usg, "EDD (CUA)", _an_date(data.get("edd_cua")))]

def build_anomaly_docx_bytes(data):
    C, J = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.JUSTIFY
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.top_margin = sec.bottom_margin = Cm(2.54)     # stationery: clinic/doctor name on top, doctor name + signature below
    sec.left_margin = sec.right_margin = Cm(1.9)
    sec.header_distance = sec.footer_distance = Cm(0)
    st_ = doc.styles["Normal"]
    st_.font.name, st_.font.size = AN_FONT, Pt(9.5)
    st_.paragraph_format.space_after = st_.paragraph_format.space_before = Pt(0)

    # ---- page 1 top spacer: 2.5 lines of 10 pt (two 10 pt + one 5 pt blank line) ----
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=5)

    # ---- title bar + patient details (details box has no blue) ----
    t = _an_table(doc, 3, [8.6, 8.6])
    title = t.cell(0, 0).merge(t.cell(0, 1))
    _an_cell(title, "USG OBS (ANOMALY SCAN)", bold=True, color="FFFFFF", fill=AN_NAVY, align=C, size=12)
    age = str(data.get("patient_age", "") or "").strip()
    agesex = f"{age}Y/{data.get('p_sex', 'F')}" if age else ""
    _an_cell(t.cell(1, 0), f"NAME :-  {data.get('patient_name', '')}", bold=True)
    _an_cell(t.cell(1, 1), f"DATE :-  {data.get('report_date', '')}", bold=True)
    _an_cell(t.cell(2, 0), f"AGE/SEX :-  {agesex}", bold=True)
    _an_cell(t.cell(2, 1), f"REF. BY :-  {data.get('referred_by', '')}", bold=True)
    _an_gap(doc)

    # ---- EGA / EDD ----
    # Same Dating / LMP / GA / EDD layout as the T3 and T16 reports (rows come from the shared dating_rows())
    C_ = WD_ALIGN_PARAGRAPH.CENTER
    _rows = dating_rows(data)
    t = _an_table(doc, 1 + len(_rows), [3.0, 4.4, 3.2, 3.2, 3.4])
    for ci, head in {0: "Dating", 1: "LMP", 4: "EDD"}.items():
        _an_cell(t.cell(0, ci), head, bold=True, color=AN_NAVY, fill=AN_LIGHT, align=C_)
    _an_cell(t.cell(0, 2).merge(t.cell(0, 3)), "GA", bold=True, color=AN_NAVY, fill=AN_LIGHT, align=C_)
    for ri, vals in enumerate(_rows, start=1):
        for ci, txt in enumerate(vals):
            _an_cell(t.cell(ri, ci), txt, bold=(ci == 0), color=(AN_NAVY if ci == 0 else None), align=C_)
    p = doc.add_paragraph()
    _an_flat(p, before=5, line=12.6)
    _an_run(p, "There is a single, live, normal intrauterine gestational pregnancy.")

    # ---- foetal survey (defaults pre-filled; placenta + numbers come from the form) ----
    _an_heading(doc, "FOETAL SURVEY & ENVIRONMENT")
    liq, liq_abn = anomaly_liquor(data)
    survey = [("FHR", f"{str(data.get('fhr', '') or '').strip() or '____'} bpm - Regular", False),
              ("Presentation", anomaly_presentation_row(data), False),
              ("Placenta", anomaly_placenta_text(data), False),
              ("Liquor", liq, liq_abn),
              ("E.F.W.", anomaly_efw(data), False),
              ("Cervix", anomaly_cervix(data), False)]
    t = _an_table(doc, len(survey), [3.6, 13.6])
    for i, (label, val, bold) in enumerate(survey):
        _an_label(t.rows[i].cells[0], label)
        _an_cell(t.rows[i].cells[1], val, bold=bold)

    _an_heading(doc, "CRANIUM")
    _an_measure_table(doc, AN_CRANIUM, data)
    _an_gap(doc)
    _an_lines(doc, AN_CRANIUM_TEXT)
    _an_heading(doc, "FOETAL SPINE")
    _an_lines(doc, AN_SPINE_TEXT)
    _an_heading(doc, "FACE")
    _an_lines(doc, AN_FACE_TEXT)
    _an_heading(doc, "ABDOMEN")
    _an_measure_table(doc, AN_ABDOMEN, data)
    _an_gap(doc)
    _an_lines(doc, AN_ABDOMEN_TEXT)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _an_flat(p, before=5)
    _an_run(p, "P.T.O", bold=True, color=AN_NAVY)

    # ---- page 2 top spacer: same 2.5 lines of 10 pt ----
    _an_spacer(doc, pt=10, page_break=True)
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=5)

    _an_heading(doc, "THORAX")
    _an_lines(doc, AN_THORAX_TEXT)
    _an_heading(doc, "UPPER EXTREMITIES")
    _an_measure_table(doc, AN_UPPER, data)
    _an_heading(doc, "LOWER EXTREMITIES")
    _an_measure_table(doc, AN_LOWER, data)
    _an_gap(doc)
    _an_lines(doc, ["No obvious abnormality is noted in the extremities."])

    _an_heading(doc, "FOETAL INDICES")
    t = _an_table(doc, len(AN_INDICES), [6.0, 4.0, 7.2])
    for i, (label, k, rng) in enumerate(AN_INDICES):
        c = t.rows[i].cells
        _an_label(c[0], label)
        _an_cell(c[1], str(data.get(k, "") or ""), align=C)
        _an_cell(c[2], rng, align=C)

    _an_heading(doc, "DOPPLER INDICES")
    drows = anomaly_doppler_rows(data)
    t = _an_table(doc, len(drows), [5.6, 1.9, 5.2, 4.5])
    for i, (label, val, pct, remark, abn) in enumerate(drows):
        c = t.rows[i].cells
        _an_label(c[0], label)
        _an_cell(c[1], str(val or ""), align=C)
        _an_pct(c[2], pct)
        _an_cell(c[3], remark, bold=abn, align=C)

    # ---- two blank lines, then the opinion ----
    for _ in range(2):
        _an_spacer(doc, pt=9.5)
    _an_heading(doc, "OPINION")
    p = doc.add_paragraph()
    _an_flat(p, after=3)
    _an_run(p, "Ultrasound Anomaly scan reveals:", bold=True, italic=True, size=7.5, color=AN_NAVY)
    p = doc.add_paragraph()
    _an_flat(p, after=7)
    _an_run(p, anomaly_opinion(data), bold=True)
    adv = anomaly_advice(data)
    if adv:                                   # pathological Doppler: advice goes on the next line, bold italics
        p = doc.add_paragraph()
        _an_flat(p, after=7)
        p.alignment = J
        _an_run(p, adv, bold=True, italic=True)

    t = _an_table(doc, 1, [17.2])
    c = t.cell(0, 0)
    set_cell_background(c, AN_LIGHT)
    c.paragraphs[0].alignment = J
    _an_flat(c.paragraphs[0])
    _an_run(c.paragraphs[0], AN_PNDT, bold=True, italic=True, size=8.5, color=AN_NAVY)

    finish_document(doc)
    normalize_table_props(doc)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def render_anomaly_preview(data):
    L = ["            USG OBS (ANOMALY SCAN)", "",
         f"NAME :- {data.get('patient_name', '')}        DATE :- {data.get('report_date', '')}",
         f"AGE/SEX :- {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}        REF. BY :- {data.get('referred_by', '')}", ""]
    L += dating_preview(data)
    L += ["", "FOETAL SURVEY & ENVIRONMENT",
          f"  FHR           : {str(data.get('fhr', '') or '').strip() or '____'} bpm - Regular",
          f"  Presentation  : {anomaly_presentation_row(data)}",
          f"  Placenta      : {anomaly_placenta_text(data)}",
          f"  Liquor        : {anomaly_liquor(data)[0]}",
          f"  E.F.W.        : {anomaly_efw(data)}",
          f"  Cervix        : {anomaly_cervix(data)}", ""]

    def block(title, spec):
        L.append(title)
        for label, k, has_pct in spec:
            row = f"  {label:<26}{_an_mm(data.get(k + '_size', '')):<8}{_an_weeks(data.get(k + '_wks', '')):<10}{_an_days(data.get(k + '_dys', '')):<9}"
            if has_pct and data.get(k + "_pct"):
                row += f"{data.get(k + '_pct')}%"
            L.append(row.rstrip())
        L.append("")
    block("CRANIUM", AN_CRANIUM); block("ABDOMEN", AN_ABDOMEN)
    block("UPPER EXTREMITIES", AN_UPPER); block("LOWER EXTREMITIES", AN_LOWER)
    L.append("FOETAL INDICES")
    for label, k, rng in AN_INDICES:
        L.append(f"  {label:<18}{str(data.get(k, '') or ''):<10}{rng}")
    L += ["", "DOPPLER INDICES"]
    for label, val, pct, remark, _ in anomaly_doppler_rows(data):
        L.append(f"  {label:<26}{str(val or ''):<7}{(str(pct) + '%') if pct else '':<9}{remark}")
    L += ["", "OPINION", "  Ultrasound Anomaly scan reveals:", f"  {anomaly_opinion(data)}"]
    if anomaly_advice(data):
        L.append(f"  {anomaly_advice(data)}")
    return "\n".join(L)


# ============================================================
# LIVE PREVIEW RENDERER
# ============================================================
def render_ob_preview(data, template_id):
    if template_id == ANOMALY:
        return render_anomaly_preview(data)
    lines = []
    lines.append(f"NAME :- {data.get('patient_name', '')}        DATE :- {data.get('report_date', '')}")
    lines.append(f"AGE/SEX :- {data.get('patient_age', '')}Y/ {data.get('p_sex', 'F')}        REF. BY:- {data.get('referred_by', '')}")
    lines.append("")
    lines.append("       COLOR DOPPLER OBSTETRICS REPORT" if template_id == COLOR_DOPPLER else "         ULTRASOUND OBSTETRICS REPORT")
    lines.append("")

    if template_id == "EARLY_OBS":
        subchorionic = "Absent" if not data.get("subchorionic_present") else f"Present; {data.get('subchorionic_grade', 'Grade-I')}"
        corpus = "Not Appreciated" if not data.get("corpus_present") else f"Appreciated; {data.get('corpus_size', '')}mm in {data.get('corpus_side', 'Right')} Ovary"
        cardiac = f"Present, {data.get('fhr_val', '')} bpm" if data.get('cardiac_status') == "Present" else data.get('cardiac_status')
        early_rows = [
            ("G-Sac", data.get('gsac', 'Seen')), ("CRL", data.get('crl', '')),
            ("Yolk sac", data.get('yolk', 'Not Seen')), ("Cardiac pulsations", cardiac),
            ("Decidual reaction", "Present"), ("Subchorionic collection", subchorionic),
            ("Corpus leuteal cyst", corpus), ("Cervix", fmt_cm(data.get('cervix', 'Normal'))),
            ("Internal Os", "Closed"), ("Placenta", "-"),
        ]
        for label, val in early_rows: lines.append(f"  {label:<28}: {val}")
        lines.append("")
        lines.extend(dating_preview(data))

    elif template_id == "NT_NB":
        lines.append("NT / NB SCAN DETAILS:")
        lines.append(f"  CRL: {data.get('crl', '')}mm")
        lines.append(f"  NT: {data.get('nt', '')}mm")
        lines.append(f"  Nasal Bone: {data.get('nasal_bone', '')}")
        lines.append(f"  FHR: {data.get('fhr_val', '')} bpm")
        lines.append(f"  DV PI: {data.get('dv_pi', '')}"
                     + (f" (Percentile: {data.get('dv_pct')}%)" if data.get("dv_pct") else "")
                     + f" (A-wave: {data.get('dv_awave', '')})")
        lines.append("")
        lines.append(f"  Placenta: {data.get('placenta_pos', '')}")
        lines.append(f"  Cervix: {fmt_cm(data.get('cervix', ''))}")
        lines.append("")
        lines.extend(dating_preview(data))

    else:
        # Combine Presentation and Lie
        pres_text = presentation_text(data)
            
        lines.append(f"Uterus shows single live fetus. Fetus is in {pres_text} at the time of scanning. Fetal movements are present. Fetal cardiac pulsations are appreciated.")
        lines.append("")
        lines.append("BIO PHYSICAL PARAMETERS:")
        for param in ["bpd", "hc", "ac", "fl", "hl"]:
            size = data.get(f"{param}_size", ""); wks = data.get(f"{param}_wks", ""); dys = data.get(f"{param}_dys", "")
            lines.append(f"  {param.upper():<4}: {(str(size) + 'mm') if size else '':<10} {(str(wks) + 'weeks') if wks else '':<12} {(str(dys) + 'days±2w') if dys else ''}")
        lines.append("")
        # --- DYNAMIC RATIO PREVIEW BASED ON GESTATIONAL AGE ---
        import re
        _mean_age_str = str(data.get('mean_age', '')).lower()
        
        # Find the digits immediately before "weeks"
        _match = re.search(r'(\d+)\s*weeks', _mean_age_str)
        
        if _match:
            _ga_weeks = int(_match.group(1))
        else:
            _ga_weeks = 0
            
        # If GA is 22 or less, use FL/HC & HC/AC. Otherwise (or if unknown), use FL/AC & FL/BPD
        if 0 < _ga_weeks <= 22:
            lines.append(f"FL/HC= {data.get('fl_hc', '')}        HC/AC = {data.get('hc_ac', '')} (BOTH WNL)")
        else:
            lines.append(f"FL/AC= {data.get('fl_ac', '')}        FL/BPD = {data.get('fl_bpd', '')} (BOTH WNL)")
        lines.append("")
        lines.append(f"MEAN AGE : {data.get('mean_age', '')}")
        
        # FETAL WEIGHT PREVIEW
        fw_str_prev = f"{data.get('fetal_weight', '')}g"
        try:
            fw_val_prev = float(data.get('fetal_weight', ''))
            err_prev = int(round(fw_val_prev * 0.15))
            fw_str_prev = f"{int(fw_val_prev)} grams ± {err_prev}gms"
        except:
            pass
            
        fw_line = f"FETAL WEIGHT : {fw_str_prev}"
        if data.get("fetal_weight_pct"):
            fw_line += f" (Percentile: {data.get('fetal_weight_pct')}%)"
        lines.append(fw_line)

        afi_method = data.get('afi_method', 'Total 4 Quadrants')
        if afi_method == "Total 4 Quadrants":
            afi_val = data.get('afi_total', '')
            desc, _ = generate_afi_text(afi_val, is_total=True)
            if desc == "Nil, Not a single pocket of even one cm seen":
                afi_str = desc
            else:
                afi_str = f"{desc}, total of four quadrants measures {afi_val} cm"
        else:
            afi_val = data.get('afi_pocket', '')
            try:
                afi_float = float(afi_val)
                if 2.0 <= afi_float <= 8.0:
                    afi_str = f"Adequate, largest pocket measures {afi_val} cm"
                elif afi_float < 2.0:
                    afi_str = f"Reduced, largest pocket measures {afi_val} cm"
                else:
                    afi_str = f"Increased, largest pocket measures {afi_val} cm"
            except:
                afi_str = f"largest pocket measures {afi_val} cm"
        lines.append(f"AFI:- {afi_str}")
        lines.append(f"Foetal heart rate is approx. {data.get('fhr', '')}.")
        lines.append("")

        lines.extend(dating_preview(data))
        lines.append("")
        plac_pos = data.get('placenta_pos', 'Anterior'); plac_os = data.get('placenta_os', 'away from the Os')
        
        # Build the base position string
        _pos_str = plac_pos.lower()
        
        if plac_os == "Placenta Praevia": 
            plac_str = f"is {_pos_str} and Placenta Praevia ({data.get('praevia_grade', 'Grade I')})."
        elif plac_os == "Low-lying":
            dist = data.get('placenta_dist', '')
            plac_str = f"is {_pos_str} and Low-lying, about {dist}cm away from the Internal Os." if dist.strip() else f"is {_pos_str} and Low-lying, but away from the Internal Os."
        else: 
            plac_str = f"is {_pos_str} and {plac_os}."
        lines.append(f"Placenta {plac_str}")

        _cx = fmt_cm(data.get('cervix', ''))
        lines.append(f"CERVIX- normal ({_cx}), internal Os is {data.get('internal_os', 'closed').lower()}." if _cx
                     else f"CERVIX- normal, internal Os is {data.get('internal_os', 'closed').lower()}.")
        if data.get('cord_around') or data.get('cord_straddling'):
            parts = []
            if data.get('cord_around'): parts.append("Single loop of cord around the neck.")
            if data.get('cord_straddling'): parts.append("Single loop of cord straddling the neck.")
            lines.append(" ".join(parts))
        else: lines.append("NO LOOP OF CORD SEEN AROUND THE NECK.")

        if template_id in ("T3", COLOR_DOPPLER) and data.get('measure_scar'): 
            lines.append(f"Scar Thickness: {scar_text(data.get('scar_size', ''))}")

        # --- DOPPLER SECTION PREVIEW ---
        if any(data.get(k) for k in DOPPLER_KEYS):
            lines.append("")
            lines.append("DOPPLER INDICES:")
            
            def _dop_line(label, prefix, extra=""):
                parts = []
                for m, name in (("pi", "PI"), ("ri", "RI"), ("sd", "S/D")):
                    if data.get(f"{prefix}_{m}"):
                        parts.append(f"{name} = {data.get(f'{prefix}_{m}')}")
                if parts:
                    lines.append(f"  {label}: " + ", ".join(parts) + extra)
            
            _dop_line("Umbilical Artery", "umbilical",
                      f" (Percentile: {data.get('umbilical_pct')}%)" if data.get("umbilical_pct") else "")
            _has_l = any(data.get(f"uterine_left_{m}") for m in ("pi", "ri", "sd"))
            _has_r = any(data.get(f"uterine_right_{m}") for m in ("pi", "ri", "sd"))
            _notch = data.get("uterine_notching", "No notching")
            _upct = f", Percentile: {data.get('uterine_pct')}%" if data.get("uterine_pct") else ""
            if _has_l and _has_r:
                _dop_line("Lt. Uterine Artery", "uterine_left")
                _dop_line("Rt. Uterine Artery", "uterine_right")
                _dop_line("Uterine Artery (Mean)", "uterine", f" ({_notch}{_upct})")
            elif _has_l:
                _dop_line("Lt. Uterine Artery", "uterine_left", f" ({_notch}{_upct})")
            elif _has_r:
                _dop_line("Rt. Uterine Artery", "uterine_right", f" ({_notch}{_upct})")
            else:
                _dop_line("Uterine Artery", "uterine", f" ({_notch}{_upct})")
            _dop_line("MCA", "mca", f" (Percentile: {data.get('mca_pct')}%)" if data.get("mca_pct") else "")
            if data.get("mca_psv"):
                _pm = f" ({data.get('mca_psv_mom')} MoM)" if data.get("mca_psv_mom") else ""
                lines.append(f"  MCA PSV: {data.get('mca_psv')} cm/s{_pm}")
            if data.get("cpr"):
                _cp = f" (Percentile: {data.get('cpr_pct')}%{'' if CPR_PERCENTILE_VALIDATED else ' - provisional'})" if data.get("cpr_pct") else ""
                lines.append(f"  CPR (MCA PI / UA PI): {data.get('cpr')}{_cp}")

    lines.append(""); lines.append("IMPRESSION:")
    mean_age = data.get('mean_age', '')
    
    imp_map = {
        "Blighted": f"INTRA UTERINE EMPTY GESTATIONAL SAC OF {data.get('ga_crl', '')} +/-1w WITH NO APPRECIABLE FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY FAILURE (BLIGHTED OVUM).",
        "NotYet": f"Intra uterine pregnancy of {data.get('ga_crl', '')} +/-1w with not yet appreciable cardiac activity. Adv- Review scan after 2 weeks.",
        "Present": f"Intra uterine live pregnancy of {data.get('ga_crl', '')} +/-1w with good cardiac activity.",
        "NT_NB": f"Single live fetus of approx. {mean_age} gestation age. NT measures {data.get('nt', '')}mm. Nasal bone is {data.get('nasal_bone', '').lower()}."
    }

    if template_id == "EARLY_OBS":
        status = data.get('cardiac_status')
        if status == "Blighted Ovum": imp_text = imp_map["Blighted"]
        elif status == "Not yet appreciable": imp_text = imp_map["NotYet"]
        else: imp_text = imp_map["Present"]
    elif template_id == "NT_NB":
        imp_text = imp_map["NT_NB"]
    else:
        pres_text = presentation_text(data)
            
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {pres_text}{',' if _is_lie_presentation(data) else ''} of approx. {mean_age} gestation age."
        
    lines.append(f"  - {imp_text}")
    if template_id == COLOR_DOPPLER and doppler_status(data)["all_normal"]:
        lines.append("")
        lines.append(f"  - {NORMAL_DOPPLER_LINE}")
    _sf = smaller_fetus_remark(data) if template_id not in ("EARLY_OBS", "NT_NB") else ""
    if _sf:
        lines.append(""); lines.append(f"  - {_sf}")
    _adv = doppler_advice(data) if template_id in ("T3", COLOR_DOPPLER) else ""
    if _adv:
        lines.append("")
        lines.append(f"  {_adv}")

    lines.append("")
    lines.append("PNDT DISCLAIMER: The radiologist has not conducted or caused to conduct any pre-natal diagnostic technique for the purpose of determining fetal sex.")
    return "\n".join(lines)
# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="OB Reporting Software", layout="wide")

st.markdown("""
    <style>
    :root, html, body { color-scheme: light !important; background-color: #f4f6f9 !important; }
    #MainMenu {visibility: hidden;} header[data-testid="stHeader"] {visibility: hidden; height: 0;}
    footer {visibility: hidden;} [data-testid="stToolbar"] {visibility: hidden;}
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stAppViewContainer"] > .main, section.main, .main .block-container { background-color: #f4f6f9 !important; }
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp li, .stApp small, .stApp strong, .stApp em { color: #111111 !important; }
    .clinic-title { text-align: center; color: #1F4E79 !important; font-weight: 700; font-size: 2rem; margin-bottom: 0; padding-bottom: 0; letter-spacing: 0.5px; }
    .clinic-sub { text-align: center; color: #555 !important; font-style: italic; font-size: 1.05rem; margin-top: 0; padding-top: 0; margin-bottom: 1.5rem; }
    .dicom-box { background-color: #f0f5fa !important; border: 1px solid #b0c4de !important; border-radius: 6px; padding: 10px; margin-bottom: 15px; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] { display: flex !important; flex-direction: column !important; justify-content: flex-end !important; }
    div[data-testid="stRadio"] label, div[data-testid="stCheckbox"] label { cursor: pointer !important; }
    div[data-testid="stExpander"] { background-color: #ffffff !important; border: 1px solid #cfd6dd !important; border-radius: 6px !important; margin-bottom: 4px !important; }
    div[data-testid="stExpander"] details, div[data-testid="stExpander"] details > div { background-color: #ffffff !important; }
    div[data-testid="stExpander"] details > summary { padding: 8px 12px !important; min-height: 42px !important; box-sizing: border-box !important; display: flex !important; align-items: center !important; background-color: #ffffff !important; }
    div[data-testid="stExpander"] details[open] > summary { background-color: #eef2f7 !important; border-bottom: 1px solid #cfd6dd !important; }
    div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary * { color: #111111 !important; font-weight: 600 !important; font-size: 14px !important; white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important; }
    div[data-testid="stNumberInput"] input { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; border-radius: 6px !important; min-height: 42px !important; box-sizing: border-box !important; font-weight: 600 !important; text-align: center !important; padding: 0 6px !important; }
    div[data-testid="stNumberInput"] button { display: none !important; }
    .stTextInput input, .stTextArea textarea, div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea, [data-baseweb="input"] input, [data-baseweb="base-input"] input, [data-baseweb="textarea"] textarea { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; min-height: 42px !important; box-sizing: border-box !important; }
    div[data-baseweb="select"] > div, div[data-baseweb="popover"] > div, ul[data-baseweb="menu"] { background-color: #ffffff !important; color: #111111 !important; }
    .stButton button, .stDownloadButton button, div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button { background-color: #ffffff !important; color: #111111 !important; border: 1px solid #cfd6dd !important; }
    .stButton button:hover, .stDownloadButton button:hover { background-color: #eef2f7 !important; border-color: #305496 !important; }
    div[data-testid="stAlert"] { background-color: #ffffff !important; color: #111111 !important; }
    div[data-testid="stCaptionContainer"], div[data-testid="stCaptionContainer"] * { color: #555555 !important; }
    hr { border-color: #cfd6dd !important; }
    div[data-testid="stMarkdownContainer"] * { color: #111111; }
    .preview-box, .stApp .preview-box, .stApp .preview-box * { color: #ffffff !important; }
    .preview-box { background-color: #2a2a2a !important; font-family: Consolas, Menlo, 'Courier New', monospace !important; font-size: 12.5px !important; line-height: 1.5 !important; padding: 14px 16px !important; border-radius: 8px !important; border: 1px solid #333 !important; white-space: pre-wrap !important; word-wrap: break-word !important; overflow-wrap: anywhere !important; word-break: break-word !important; margin-bottom: 8px !important; display: block !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="clinic-title">डॉ. प्रणव एवं डॉ. गौरव इमेजिंग, आगरा</div>', unsafe_allow_html=True)
st.markdown('<div class="clinic-sub">Precision Imaging. Trusted Diagnostics. Agra.</div>', unsafe_allow_html=True)

# --- NEW: CHECK LISTENER JSON AT THE VERY TOP ---
import json
import os
json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'latest_sr.json')

if os.path.exists(json_path):
    mtime = None
    try:
        mtime = os.path.getmtime(json_path)
        if st.session_state.get("last_json_mtime") != mtime and st.session_state.get("json_bad_mtime") != mtime:
            with open(json_path, 'r') as f:
                new_data = json.load(f)
            _old_v = st.session_state.get("dicom_version", 0)
            _had_data = bool(st.session_state.get("sr_data"))
            st.session_state["sr_data"] = new_data
            st.session_state["last_json_mtime"] = mtime
            st.session_state["dicom_version"] = _old_v + 1
            if not _had_data:
                # The form was empty, so anything typed before the SR arrived (e.g. the DV PI) belongs to this exam:
                # move it to the new form version. Values the SR supplies overwrite it below; the rest stays.
                # (If a previous patient's data was still loaded nothing is carried over - no cross-patient leakage.)
                _sfx = f"_{_old_v}"
                for _k in list(st.session_state.keys()):
                    if _k.startswith("sr_") and _k.endswith(_sfx):
                        st.session_state[_k[:-len(_sfx)] + f"_{_old_v + 1}"] = st.session_state[_k]
            st.toast("✅ Loaded fresh data from DICOM Listener!")
    except Exception as e:
        st.session_state["json_bad_mtime"] = mtime      # unreadable right now: do not retry until the file changes
# -----------------------------------------------

v = st.session_state.get("dicom_version", 0)
sr_data = st.session_state.get("sr_data", {})

# --- AUTO-REFRESH: the moment the listener writes a new latest_sr.json, reload the app (checks every 2 s) ---
def _watch_listener_file():
    try:
        if os.path.exists(json_path):
            _m = os.path.getmtime(json_path)
            if _m != st.session_state.get("last_json_mtime") and _m != st.session_state.get("json_bad_mtime"):
                st.rerun(scope="app")
    except FileNotFoundError:
        pass
if hasattr(st, "fragment"):
    st.fragment(run_every=2)(_watch_listener_file)()
else:
    st.caption("Update Streamlit (pip install -U streamlit) to refresh automatically when a new SR arrives.")

# --- FIX FOR PLACENTA DROPDOWN CACHING ---
if "placenta_pos_widget" not in st.session_state:
    st.session_state["placenta_pos_widget"] = "Anterior"
if "placenta_os_widget" not in st.session_state:
    st.session_state["placenta_os_widget"] = "away from the Os"

# --- THE ACTUAL NAME/AUTO-FILL FIX ---
# We bypass the versioning and force the name/age directly into the widget keys.
if 'patient_name' in sr_data and sr_data['patient_name']:
    st.session_state["p_name_input"] = sr_data['patient_name']
if 'patient_age' in sr_data and sr_data['patient_age']:
    st.session_state["p_age_input"] = sr_data['patient_age']

c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3])
with c1: p_name = st.text_input("Name", key="p_name_input")
with c2: p_age = st.text_input("Age", key="p_age_input")
with c3: p_sex = st.radio("Sex", ["F", "M"], horizontal=True)
with c4: p_date = st.text_input("Date", value=datetime.now().strftime("%d-%b-%y"))
with c5: p_ref = st.text_input("Referred by")

data = {"patient_name": p_name, "patient_age": p_age, "report_date": p_date, "referred_by": p_ref, "p_sex": p_sex}

# --- NEW PATIENT BUTTON ---
if st.button("🚀 Start New Patient / Clear Form", type="primary"):
    # Clear the DICOM data
    st.session_state["sr_data"] = {}
    st.session_state["dicom_version"] += 1
    # Clear the input fields by deleting their session state keys
    for key in list(st.session_state.keys()):
        if key.startswith("sr_") or key in ["p_name_input", "p_age_input"]:
            del st.session_state[key]
    st.rerun()
# ---------------------------

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    template_id = st.selectbox("Template", ["EARLY_OBS", "NT_NB", "T16", "T3", COLOR_DOPPLER, ANOMALY])
    
    # --- DUAL DICOM INPUT (Manual Upload + Listener) ---
    import json
    import os
    
    st.markdown("### 🔗 DICOM Input")
    
    # Initialize the version bumper
    if "dicom_version" not in st.session_state:
        st.session_state["dicom_version"] = 0
        
    col_up, col_ref = st.columns([2, 1])
    
    with col_up:
        uploaded_sr = st.file_uploader("Auto-Fill from File", type=["dcm", "dicom", "ima", "json"], label_visibility="collapsed")
        if uploaded_sr is None:
            st.session_state["last_upload_id"] = None
        else:
            # A file sitting in the uploader is applied ONCE. It used to be re-applied on every rerun, which bumped the
            # form version each time - so everything typed by hand (e.g. the DV PI) was wiped as soon as you hit Enter.
            _up_id = (uploaded_sr.name, getattr(uploaded_sr, "size", None))
            if st.session_state.get("last_upload_id") != _up_id:
                try:
                    if uploaded_sr.name.endswith('.json'):
                        st.session_state["sr_data"] = json.load(uploaded_sr)
                    else:
                        ds = pydicom.dcmread(io.BytesIO(uploaded_sr.read()))
                        st.session_state["sr_data"] = extract_sr_data(ds)

                    st.session_state["dicom_version"] += 1  # Force refresh
                    # Mark the listener file as already seen so it does not overwrite this upload on the next rerun
                    st.session_state["last_json_mtime"] = os.path.getmtime(json_path) if os.path.exists(json_path) else None
                    st.session_state["last_upload_id"] = _up_id
                    st.success("DICOM SR File Loaded (Manual)")
                except Exception as e:
                    st.error(f"Error reading SR: {e}")
            else:
                st.caption(f"Manual file loaded: {uploaded_sr.name}")
                     
    with col_ref:
        if st.button("🔄 Refresh from Listener", use_container_width=True, key="refresh_btn"):
            st.session_state["last_json_mtime"] = None  # Forces the app to look at the file again
            st.rerun()
            
    # Push values into session_state widgets with versioning
    sr_data = st.session_state.get("sr_data", {})
    v = st.session_state["dicom_version"]
    
    if "last_applied_version" not in st.session_state:
        st.session_state["last_applied_version"] = -1
        
    apply_new_data = st.session_state["last_applied_version"] != v
    
    if apply_new_data:
        # --- NEW: Auto-fill Patient Info ---
        if 'patient_name' in sr_data: st.session_state[f"sr_patient_name_{v}"] = sr_data['patient_name']
        if 'patient_age' in sr_data: st.session_state[f"sr_patient_age_{v}"] = sr_data['patient_age']
        
        if 'bpd' in sr_data: st.session_state[f"sr_bpd_size_{v}"] = sr_data['bpd']
        if 'hc' in sr_data: st.session_state[f"sr_hc_size_{v}"] = sr_data['hc']
        if 'ac' in sr_data: st.session_state[f"sr_ac_size_{v}"] = sr_data['ac']
        if 'fl' in sr_data: st.session_state[f"sr_fl_size_{v}"] = sr_data['fl']
        if 'hl' in sr_data: st.session_state[f"sr_hl_size_{v}"] = sr_data['hl']
        
        for p in ["bpd", "hc", "ac", "fl", "hl"]:
            if f"{p}_wks" in sr_data: st.session_state[f"sr_{p}_wks_{v}"] = sr_data[f"{p}_wks"]
            if f"{p}_dys" in sr_data: st.session_state[f"sr_{p}_dys_{v}"] = sr_data[f"{p}_dys"]
            if f"{p}_pct" in sr_data:
                st.session_state[f"sr_{p}_pct_{v}"] = sr_data[f"{p}_pct"]
                # The visible box uses the "_edit_" key. When an empty box was carried over from the previous form version
                # it would override the SR value (blank bars), so seed the box itself too.
                st.session_state[f"sr_{p}_pct_edit_{v}"] = sr_data[f"{p}_pct"]

        # Anomaly scan: sizes / GA / percentile straight into the Anomaly widgets
        for _k in ("tcd", "ulna", "tibia"):
            if _k in sr_data: st.session_state[f"sr_an_{_k}_size_{v}"] = sr_data[_k]
            for _s in ("wks", "dys", "pct"):
                if f"{_k}_{_s}" in sr_data: st.session_state[f"sr_an_{_k}_{_s}_{v}"] = sr_data[f"{_k}_{_s}"]
        for _k in ("lv", "cm", "nf", "radius", "fibula"):
            if _k in sr_data: st.session_state[f"sr_an_{_k}_size_{v}"] = sr_data[_k]
        if "dv_pi" in sr_data: st.session_state[f"sr_an_dv_pi_{v}"] = sr_data["dv_pi"]

        if 'msd' in sr_data: st.session_state[f"sr_msd_{v}"] = sr_data['msd']
        if 'crl' in sr_data: st.session_state[f"sr_crl_{v}"] = sr_data['crl']
        if 'nt' in sr_data: st.session_state[f"sr_nt_{v}"] = sr_data['nt']
        if 'nasal_bone' in sr_data: st.session_state[f"sr_nasal_bone_{v}"] = sr_data['nasal_bone']
        if 'fetal_weight' in sr_data: st.session_state[f"sr_fetal_weight_{v}"] = sr_data['fetal_weight']
        if 'fetal_weight_pct' in sr_data: st.session_state[f"sr_fetal_weight_pct_edit_{v}"] = sr_data['fetal_weight_pct']
        if 'fhr' in sr_data: st.session_state[f"sr_fhr_{v}"] = sr_data['fhr']
        if 'mean_age' in sr_data: st.session_state[f"sr_mean_age_{v}"] = sr_data['mean_age']
        
        if 'edd_lmp' in sr_data: st.session_state[f"sr_edd_lmp_{v}"] = sr_data['edd_lmp']
        if 'edd_cua' in sr_data: st.session_state[f"sr_edd_cua_{v}"] = sr_data['edd_cua']
        if 'lmp' in sr_data: st.session_state[f"sr_lmp_{v}"] = sr_data['lmp']
        
        if 'ga_lmp_days' in sr_data: 
            st.session_state[f"sr_ga_lmp_{v}"] = f"{sr_data['ga_lmp_days']//7}w{sr_data['ga_lmp_days']%7}d"

        # Doppler: PI / RI / S/D for each vessel (uterine has left, right and mean)
        for _vsl in ["umbilical", "uterine", "uterine_left", "uterine_right", "mca"]:
            for _m in ["pi", "ri", "sd"]:
                _k = f"{_vsl}_{_m}"
                if _k in sr_data:
                    st.session_state[f"sr_{_k}_{v}"] = sr_data[_k]
        if "mca_psv" in sr_data:
            st.session_state[f"sr_mca_psv_{v}"] = sr_data["mca_psv"]
        
        st.session_state["last_applied_version"] = v

    # --------------------------
    lmp_not_sure = st.checkbox("LMP Not Sure / Not Known")
    if lmp_not_sure:
        data["lmp"] = "Not Sure"
        data["ga_lmp"] = ""
    else:
        data["lmp"] = st.text_input("LMP", value=st.session_state.get(f"sr_lmp_{v}", ""))
        data["ga_lmp"] = st.text_input("GA by LMP", value=st.session_state.get(f"sr_ga_lmp_{v}", ""), key=f"ga_lmp_input_{v}")

    with st.expander("🔗 DICOM SR Data (Machine Generated)", expanded=False):
        st.markdown('<div class="dicom-box">Verify or override data received from the ultrasound machine.</div>', unsafe_allow_html=True)
        
        if template_id != "EARLY_OBS":
            # Timeline Header
            st.markdown(f"**LMP:** {st.session_state.get(f'sr_lmp_{v}', '')} &nbsp;|&nbsp; "
                        f"**GA(LMP):** {st.session_state.get(f'sr_ga_lmp_{v}', '')} &nbsp;|&nbsp; "
                        f"**AUA:** {st.session_state.get(f'sr_mean_age_{v}', '')} &nbsp;|&nbsp; "
                        f"**EDD(LMP):** {st.session_state.get(f'sr_edd_lmp_{v}', '')}")
            st.markdown("---")

        if template_id == "EARLY_OBS":
            data["gsac"] = st.text_input("G-Sac", value=st.session_state.get(f"sr_gsac_{v}", "Seen"), key=f"sr_gsac_{v}")
            data["msd"] = st.text_input("MSD", value=st.session_state.get(f"sr_msd_{v}", ""), key=f"sr_msd_{v}")
            data["crl"] = st.text_input("CRL", value=st.session_state.get(f"sr_crl_{v}", ""), key=f"sr_crl_{v}")
            data["ga_crl"] = st.text_input("GA by CRL", value="")
            data["yolk"] = st.text_input("Yolk sac", value="")
            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
        elif template_id == "NT_NB":
            b1, b2 = st.columns(2)
            data["crl"] = b1.text_input("CRL (mm)", value=st.session_state.get(f"sr_crl_{v}", ""), key=f"sr_crl_{v}")
            data["nt"] = b2.text_input("NT (mm)", value=st.session_state.get(f"sr_nt_{v}", ""), key=f"sr_nt_{v}")
            
            b1, b2 = st.columns(2)
            data["nasal_bone"] = b1.text_input("Nasal Bone", value=st.session_state.get(f"sr_nasal_bone_{v}", ""), key=f"sr_nasal_bone_{v}")
            data["fhr_val"] = b2.text_input("FHR (bpm)", value=st.session_state.get(f"sr_fhr_{v}", ""), key=f"sr_fhr_{v}")
            
            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
        else:
            st.markdown("**Fetal Growth Parameters**")
            h1, h2, h3, h4, h5 = st.columns([2, 1, 1, 1, 3])
            h1.caption("Parameter")
            h2.caption("Size (mm)")
            h3.caption("Wks")
            h4.caption("Days")
            h5.caption("Percentile")

            params = [
                ("bpd", "Biparietal Diameter"),
                ("hc", "Head Circumference"),
                ("ac", "Abdominal Circumference"),
                ("fl", "Femur Length"),
                ("hl", "Humerus Length")
            ]

            for p_key, p_name in params:
                c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 3])
                c1.markdown(f"**{p_name}**")
                
                data[f"{p_key}_size"] = c2.text_input(f"{p_key} size", value=st.session_state.get(f"sr_{p_key}_size_{v}", ""), key=f"sr_{p_key}_size_{v}", label_visibility="collapsed")
                data[f"{p_key}_wks"] = c3.text_input(f"{p_key} wks", value=st.session_state.get(f"sr_{p_key}_wks_{v}", ""), key=f"sr_{p_key}_wks_{v}", label_visibility="collapsed")
                data[f"{p_key}_dys"] = c4.text_input(f"{p_key} dys", value=st.session_state.get(f"sr_{p_key}_dys_{v}", ""), key=f"sr_{p_key}_dys_{v}", label_visibility="collapsed")
                
                # Make the percentile an editable text box, pre-filled with the DICOM value
                pct_val = c5.text_input(f"{p_key} pct", value=st.session_state.get(f"sr_{p_key}_pct_{v}", ""), key=f"sr_{p_key}_pct_edit_{v}", label_visibility="collapsed")
                
                if pct_val:
                    bar_img = generate_percentile_bar(pct_val)
                    if bar_img:
                        c5.image(bar_img, use_container_width=True)
            st.markdown("---")
            st.markdown("**Ratios & Averages**")
            b1, b2 = st.columns(2)
            
            try:
                if data.get("fl_size") and data.get("hc_size"):
                    fl_hc_val = (float(data["fl_size"]) / float(data["hc_size"])) * 100
                    st.session_state[f"sr_fl_hc_{v}"] = f"{fl_hc_val:.1f}%"
                if data.get("hc_size") and data.get("ac_size"):
                    hc_ac_val = float(data["hc_size"]) / float(data["ac_size"])
                    st.session_state[f"sr_hc_ac_{v}"] = f"{hc_ac_val:.2f}"
                if data.get("fl_size") and data.get("bpd_size"):
                    fl_bpd_val = (float(data["fl_size"]) / float(data["bpd_size"])) * 100
                    st.session_state[f"sr_fl_bpd_{v}"] = f"{fl_bpd_val:.1f}%"
                if data.get("fl_size") and data.get("ac_size"):
                    fl_ac_val = (float(data["fl_size"]) / float(data["ac_size"])) * 100
                    st.session_state[f"sr_fl_ac_{v}"] = f"{fl_ac_val:.1f}%"
            except:
                pass

            # --- CRITICAL FIX: Pull calculated percentiles from session_state into data dict ---
            for p in ["bpd", "hc", "ac", "fl", "hl"]:
                pct_box_val = st.session_state.get(f"sr_{p}_pct_edit_{v}", "")
                
                # If LMP is "Not Sure", we send empty data to the document/preview
                if lmp_not_sure:
                    data[f"{p}_pct"] = ""
                elif pct_box_val:
                    data[f"{p}_pct"] = pct_box_val

            data["fl_hc"] = b1.text_input("FL/HC", value=st.session_state.get(f"sr_fl_hc_{v}", ""), key=f"sr_fl_hc_{v}")
            data["hc_ac"] = b2.text_input("HC/AC", value=st.session_state.get(f"sr_hc_ac_{v}", ""), key=f"sr_hc_ac_{v}")
            
            b1, b2 = st.columns(2)
            data["fl_bpd"] = b1.text_input("FL/BPD", value=st.session_state.get(f"sr_fl_bpd_{v}", ""), key=f"sr_fl_bpd_{v}")
            data["fl_ac"] = b2.text_input("FL/AC", value=st.session_state.get(f"sr_fl_ac_{v}", ""), key=f"sr_fl_ac_{v}")

            b1, b2 = st.columns(2)
            data["mean_age"] = b1.text_input("Mean Age", value=st.session_state.get(f"sr_mean_age_{v}", ""), key=f"sr_mean_age_{v}")
            data["fetal_weight"] = b2.text_input("Fetal Weight (g)", value=st.session_state.get(f"sr_fetal_weight_{v}", ""), key=f"sr_fetal_weight_{v}")
            
            # EFW Percentile Fix
            # Streamlit doesn't update the data dictionary automatically when you change a widget with a key. 
            # We have to pull it directly from session_state right before building the doc.
            efw_pct_val = b2.text_input("EFW Percentile (%)", value=st.session_state.get(f"sr_fetal_weight_pct_{v}", ""), key=f"sr_fetal_weight_pct_edit_{v}")
            
            if lmp_not_sure:
                data["fetal_weight_pct"] = ""
            else:
                data["fetal_weight_pct"] = efw_pct_val
            if efw_pct_val:
                bar_efw = generate_percentile_bar(efw_pct_val)
                if bar_efw: b2.image(bar_efw, use_container_width=True)
            b1, b2 = st.columns(2)
            data["fhr"] = b1.text_input("FHR", value=st.session_state.get(f"sr_fhr_{v}", ""), key=f"sr_fhr_{v}")
            data["edd_lmp"] = b2.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
            
            # --- DOPPLER INPUTS ---
            def _mean_of(*vals):
                nums = []
                for x in vals:
                    try: nums.append(float(x))
                    except (TypeError, ValueError): pass
                return f"{sum(nums)/len(nums):.2f}" if nums else ""

            st.markdown("**Doppler Indices & Percentiles**")

            st.caption("Umbilical Artery")
            c1, c2, c3 = st.columns(3)
            data["umbilical_pi"] = c1.text_input("UA PI", key=f"sr_umbilical_pi_{v}")
            data["umbilical_ri"] = c2.text_input("UA RI", key=f"sr_umbilical_ri_{v}")
            data["umbilical_sd"] = c3.text_input("UA S/D", key=f"sr_umbilical_sd_{v}")

            st.caption("Uterine Arteries")
            c1, c2, c3 = st.columns(3)
            data["uterine_left_pi"] = c1.text_input("Left PI", key=f"sr_uterine_left_pi_{v}")
            data["uterine_left_ri"] = c2.text_input("Left RI", key=f"sr_uterine_left_ri_{v}")
            data["uterine_left_sd"] = c3.text_input("Left S/D", key=f"sr_uterine_left_sd_{v}")
            c1, c2, c3 = st.columns(3)
            data["uterine_right_pi"] = c1.text_input("Right PI", key=f"sr_uterine_right_pi_{v}")
            data["uterine_right_ri"] = c2.text_input("Right RI", key=f"sr_uterine_right_ri_{v}")
            data["uterine_right_sd"] = c3.text_input("Right S/D", key=f"sr_uterine_right_sd_{v}")

            # Mean = average of Left/Right as typed; falls back to the SR's own unsided value.
            # The mean is part of the widget key so it refreshes when Left/Right change.
            mean_pi = _mean_of(data["uterine_left_pi"], data["uterine_right_pi"]) or sr_data.get("uterine_pi", "")
            mean_ri = _mean_of(data["uterine_left_ri"], data["uterine_right_ri"]) or sr_data.get("uterine_ri", "")
            mean_sd = _mean_of(data["uterine_left_sd"], data["uterine_right_sd"]) or sr_data.get("uterine_sd", "")
            c1, c2, c3 = st.columns(3)
            data["uterine_pi"] = c1.text_input("Mean Uterine PI", value=mean_pi, key=f"sr_uterine_pi_mean_{v}_{mean_pi}")
            data["uterine_ri"] = c2.text_input("Mean Uterine RI", value=mean_ri, key=f"sr_uterine_ri_mean_{v}_{mean_ri}")
            data["uterine_sd"] = c3.text_input("Mean Uterine S/D", value=mean_sd, key=f"sr_uterine_sd_mean_{v}_{mean_sd}")

            # MCA only appears when the SR contains it
            mca_present = template_id == COLOR_DOPPLER or any(k in sr_data for k in ("mca_pi", "mca_ri", "mca_sd"))
            if mca_present:
                st.caption("Middle Cerebral Artery")
                if template_id == COLOR_DOPPLER:
                    c1, c2, c3, c4 = st.columns(4)
                else:
                    c1, c2, c3 = st.columns(3)
                    c4 = None
                data["mca_pi"] = c1.text_input("MCA PI", key=f"sr_mca_pi_{v}")
                data["mca_ri"] = c2.text_input("MCA RI", key=f"sr_mca_ri_{v}")
                data["mca_sd"] = c3.text_input("MCA S/D", key=f"sr_mca_sd_{v}")
                data["mca_psv"] = c4.text_input("MCA PSV (cm/s)", key=f"sr_mca_psv_{v}") if c4 is not None else ""
            else:
                data["mca_pi"] = data["mca_ri"] = data["mca_sd"] = data["mca_psv"] = ""  # stays blank if not present
            
            data["uterine_notching"] = st.selectbox("Uterine Artery Notching", ["No notching", "Unilateral Notching", "Bilateral Notching"])

            # Default GA for percentiles = ultrasound GA from the SR (falls back to LMP GA, then 28 w)
            _ga_days_sr = sr_data.get("ga_usg_days") or sr_data.get("ga_lmp_days")
            _ga_default = round(min(max(_ga_days_sr / 7.0, 10.0), 42.0), 1) if _ga_days_sr else 28.0
            data["ga_weeks_dop"] = st.number_input("GA (weeks) for Percentiles", min_value=10.0, max_value=42.0, value=_ga_default, step=0.1)

            b1, b2, b3 = st.columns(3)
            if data.get("umbilical_pi"):
                try:
                    res = calculate_ua_pi(float(data["umbilical_pi"]), data["ga_weeks_dop"])
                    data["umbilical_pct"] = str(res['percentile'])
                    b1.caption(f"Umbilical: {res['percentile']}th")
                    bar_ua = generate_percentile_bar(res['percentile'])
                    if bar_ua: b1.image(bar_ua, use_container_width=True)
                except: pass
                
            if data.get("uterine_pi"):
                try:
                    res = calculate_uta_pi(float(data["uterine_pi"]), data["ga_weeks_dop"])
                    if res: 
                        data["uterine_pct"] = str(res['percentile'])
                        b2.caption(f"Uterine: {res['percentile']}th")
                        bar_uta = generate_percentile_bar(res['percentile'])
                        if bar_uta: b2.image(bar_uta, use_container_width=True)
                except: pass

            if data.get("mca_pi"):
                try:
                    res = calculate_mca_pi(float(data["mca_pi"]), data["ga_weeks_dop"])
                    if res:
                        data["mca_pct"] = str(res['percentile'])
                        b3.caption(f"MCA: {res['percentile']}th")
                        bar_mca = generate_percentile_bar(res['percentile'])
                        if bar_mca: b3.image(bar_mca, use_container_width=True)
                except: pass

            if template_id == COLOR_DOPPLER:
                # Right / Left uterine percentiles: the same reference curve applied to each side
                for _side in ("left", "right"):
                    try:
                        if data.get(f"uterine_{_side}_pi"):
                            _r = calculate_uta_pi(float(data[f"uterine_{_side}_pi"]), data["ga_weeks_dop"])
                            if _r:
                                data[f"uterine_{_side}_pct"] = str(_r['percentile'])
                    except: pass
                # Cerebro-placental ratio = MCA PI / UA PI
                try:
                    if data.get("mca_pi") and data.get("umbilical_pi"):
                        _cpr = float(data["mca_pi"]) / float(data["umbilical_pi"])
                        data["cpr"] = f"{_cpr:.2f}"
                        _rc = calculate_cpr(_cpr, data["ga_weeks_dop"])
                        if _rc:
                            data["cpr_pct"] = str(_rc['percentile'])
                            c1, c2 = st.columns([1, 1])
                            c1.caption(f"CPR: {data['cpr']} ({_rc['percentile']}th percentile)")
                            _bar_cpr = generate_percentile_bar(_rc['percentile'])
                            if _bar_cpr: c2.image(_bar_cpr, use_container_width=True)
                            if not CPR_PERCENTILE_VALIDATED:
                                st.warning("CPR percentile is provisional (not yet matched to the Barcelona calculator). "
                                           "It is marked with a dagger in the report.")
                except: pass
                # MCA peak systolic velocity -> MoM
                try:
                    if data.get("mca_psv"):
                        _rp = calculate_mca_psv_mom(float(data["mca_psv"]), data["ga_weeks_dop"])
                        if _rp:
                            data["mca_psv_mom"] = f"{_rp['mom']:.2f}"
                            data["mca_psv_median"] = f"{_rp['median']:.1f}"
                            _msg = f"MCA PSV: {data['mca_psv']} cm/s = {data['mca_psv_mom']} MoM (median {data['mca_psv_median']} cm/s)"
                            if _rp["mom"] >= PSV_MOM_ABNORMAL:
                                st.warning(_msg + " - at or above 1.5 MoM")
                            else:
                                st.caption(_msg)
                except: pass

    if template_id == "EARLY_OBS":
        with st.expander("Early OBS Details", expanded=True):
            data["cardiac_status"] = st.selectbox("Cardiac Activity", ["Present", "Not yet appreciable", "Blighted Ovum"])
            if data["cardiac_status"] == "Present":
                data["fhr_val"] = st.text_input("FHR Value (bpm)", value=st.session_state.get("sr_fhr", ""))
            data["cervix"] = st.text_input("Cervix Length", value="Normal")
            data["subchorionic_present"] = st.checkbox("Subchorionic collection present")
            if data["subchorionic_present"]:
                data["subchorionic_grade"] = st.radio("Grade", ["Grade-I", "Grade-I/II", "Grade-II", "Grade-II/III"], horizontal=True)
            data["corpus_present"] = st.checkbox("Corpus Luteum Cyst present")
            if data["corpus_present"]:
                b1, b2 = st.columns(2)
                data["corpus_side"] = b1.selectbox("Side", ["Right", "Left"])
                data["corpus_size"] = b2.text_input("Size (mm)")

    elif template_id == "NT_NB":
        with st.expander("Typist Inputs (NT / NB)", expanded=True):
            data["dv_pi"] = st.text_input("Ductus Venosus PI", value="")
            data["dv_awave"] = st.selectbox("DV A-Wave", ["Positive", "Absent", "Reversed"])
            
            b1, b2 = st.columns(2)
            _plac_opts = PLACENTA_OPTS
            # The key forces Streamlit to refresh the widget
            data["placenta_pos"] = b1.selectbox("Placenta Position", _plac_opts, key="placenta_pos_widget")
            data["placenta_os"] = b2.selectbox("Placenta Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key="placenta_os_widget")

    else:
        with st.expander("Anatomy & Presentation", expanded=True):
            b1, b2 = st.columns(2)
            _pres_default = PRESENTATION_OPTIONS.index(PRESENTATION_DEFAULTS.get(template_id, "CEPHALIC"))
            data["presentation"] = b1.selectbox("Presentation", PRESENTATION_OPTIONS, index=_pres_default,
                                                key=f"presentation_widget_{template_id}")
            if data["presentation"] == "CEPHALIC":
                data["lie"] = b2.selectbox("Lie", LIE_OPTIONS, format_func=lambda c: LIE_LABELS.get(c, c), key="lie_widget")
            else:
                b2.selectbox("Lie", ["None"], disabled=True, key="lie_widget_disabled")  # lie options only apply to cephalic
                data["lie"] = "None"
            if data["presentation"] in ("OBLIQUE LIE", "TRANSVERSE LIE"):
                q1, q2 = st.columns(2)
                data["head_side"] = q1.selectbox("Head in - Side", ["Right", "Left"], key="head_side_widget")
                data["head_level"] = q2.selectbox("Head in - Quadrant", ["Upper", "Lower"], key="head_level_widget")

            data["afi_method"] = st.radio("AFI Method", ["Largest Pocket", "Total 4 Quadrants"], horizontal=True)
            if data["afi_method"] == "Total 4 Quadrants":
                data["afi_total"] = st.text_input("AFI (cm)")
            else:
                data["afi_pocket"] = st.text_input("Largest Pocket (cm)")

            b1, b2 = st.columns(2)
            data["placenta_pos"] = b1.selectbox("Placenta Position", PLACENTA_OPTS)
            data["placenta_os"] = b2.selectbox("Placenta Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"])

            if data["placenta_os"] == "Low-lying":
                data["placenta_dist"] = st.text_input("Distance from Internal Os (cm) [Optional]")
            elif data["placenta_os"] == "Placenta Praevia":
                data["praevia_grade"] = st.radio("Praevia Grade", ["Grade I", "Grade II", "Grade III", "Grade IV"], horizontal=True)

            b1, b2 = st.columns(2)
            data["cervix"] = b1.text_input("Cervix Length (cm)")
            data["internal_os"] = b2.selectbox("Internal Os", ["Closed", "Open"])

            if template_id != ANOMALY:        # cord loop is not part of the Anomaly Scan report
                st.markdown("**Cord Loop**")
                data["cord_around"] = st.checkbox("Single loop of cord around the neck")
                data["cord_straddling"] = st.checkbox("Single loop of cord straddling the neck")

            if template_id in ("T3", COLOR_DOPPLER):
                st.markdown("**Scar Thickness**")
                data["measure_scar"] = st.checkbox("Measure Scar Thickness")
                if data["measure_scar"]:
                    data["scar_size"] = st.text_input("Scar Size (mm)")

    if template_id == ANOMALY:
        with st.expander("🧠 Anomaly Scan Measurements", expanded=True):
            st.caption("Filled from the DICOM SR when the machine sends them - type anything that is missing. Percentiles typed here are manual.")
            h1, h2, h3, h4, h5 = st.columns([2, 1, 1, 1, 3])
            h1.caption("Parameter"); h2.caption("Size (mm)"); h3.caption("Wks"); h4.caption("Days"); h5.caption("Percentile")
            for _k, _label in (("tcd", "Cerebellum (TCD)"), ("ulna", "Ulna"), ("tibia", "Tibia")):
                c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 3])
                c1.markdown(f"**{_label}**")
                data[f"{_k}_size"] = c2.text_input(f"{_k} size", key=f"sr_an_{_k}_size_{v}", label_visibility="collapsed")
                data[f"{_k}_wks"] = c3.text_input(f"{_k} wks", key=f"sr_an_{_k}_wks_{v}", label_visibility="collapsed")
                data[f"{_k}_dys"] = c4.text_input(f"{_k} dys", key=f"sr_an_{_k}_dys_{v}", label_visibility="collapsed")
                data[f"{_k}_pct"] = c5.text_input(f"{_k} pct", key=f"sr_an_{_k}_pct_{v}", label_visibility="collapsed")
                if data[f"{_k}_pct"]:
                    _bar = generate_percentile_bar(data[f"{_k}_pct"])
                    if _bar: c5.image(_bar, use_container_width=True)
            st.markdown("**Size only (mm)**")
            s1, s2, s3, s4, s5 = st.columns(5)
            data["lv_size"] = s1.text_input("Lateral Ventricle", key=f"sr_an_lv_size_{v}")
            data["cm_size"] = s2.text_input("Cisterna Magna", key=f"sr_an_cm_size_{v}")
            data["nf_size"] = s3.text_input("Nuchal Fold", key=f"sr_an_nf_size_{v}")
            data["radius_size"] = s4.text_input("Radius", key=f"sr_an_radius_size_{v}")
            data["fibula_size"] = s5.text_input("Fibula", key=f"sr_an_fibula_size_{v}")
            data["efw_sd"] = st.text_input("EFW \u00b1 (gms)  [leave blank for automatic 15%]", key=f"sr_an_efw_sd_{v}")
            st.markdown("**Ductus Venosus** (type the PI - the percentile and bar fill in automatically)")
            d1, d2, d3 = st.columns(3)
            data["dv_pi"] = d1.text_input("DV PI", key=f"sr_an_dv_pi_{v}")
            data["dv_awave"] = d3.selectbox("DV A-Wave", ["Positive", "Absent", "Reversed"], key=f"sr_an_dv_awave_{v}")
            _dv_auto = ""
            try:
                if data["dv_pi"].strip():
                    _rdv = calculate_dv_pi(float(data["dv_pi"].replace(",", ".")), float(data.get("ga_weeks_dop") or 0))
                    if _rdv:
                        _dv_auto = str(_rdv["percentile"])
                    else:
                        d1.caption("No percentile outside 20-40 weeks (Barcelona range) - type one if needed.")
            except (ValueError, TypeError):
                d1.caption("DV PI must be a number.")
            # The box is filled with the calculated percentile and re-fills whenever the PI or GA changes (the value
            # is part of the key); typing in it overrides the calculation until the PI or GA changes again.
            data["dv_pct"] = d2.text_input("DV PI percentile (auto)", value=_dv_auto,
                                           key=f"sr_an_dv_pct_{v}_{_dv_auto}").strip()
            if data["dv_pct"]:
                _bar = generate_percentile_bar(data["dv_pct"])
                if _bar: d2.image(_bar, use_container_width=True)


with col2:
    st.subheader("📄 Live Preview")
    _preview_text = render_ob_preview(data, template_id)
    _safe = html.escape(_preview_text).replace("\n", "<br>")
    st.markdown(f'<div class="preview-box">{_safe}</div>', unsafe_allow_html=True)

    st.subheader("⬇️ Download Report")
    _is_doppler_tpl = template_id == COLOR_DOPPLER
    _is_anomaly_tpl = template_id == ANOMALY
    _color_idx = PINK_THEME_INDEX if _is_doppler_tpl else load_color_index()   # Color Doppler is always pink
    data["color_theme"] = _color_idx
    docx_bytes = build_docx_bytes(data, template_id)
    
    # Generate a unique filename: PatientName_GA_DateTime
    import datetime as dt
    _now = dt.datetime.now().strftime("%d%b%Y_%H%M")
    
    # Clean up the name and GA to avoid invalid file characters
    _clean_name = str(data.get('patient_name', 'Unknown')).replace(" ", "_").replace("/", "_")
    _clean_ga = str(data.get('mean_age', 'UnknownGA')).replace(" ", "").replace("±", "")
    
    _filename = f"{'DOPPLER' if _is_doppler_tpl else 'ANOMALY' if _is_anomaly_tpl else 'OB'}_{_clean_name}_{_clean_ga}_{_now}.docx"

    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        on_click=None if (_is_doppler_tpl or _is_anomaly_tpl) else advance_color_index   # Color Doppler does not use up a rotation colour
    )
    if _is_anomaly_tpl:
        st.caption("Table colour: Blue (fixed for Anomaly Scan reports; the colour rotation is not affected)")
    elif _is_doppler_tpl:
        st.caption("Table colour: Pink (fixed for Color Doppler reports; the colour rotation is not affected)")
    else:
        st.caption(f"Table colour for this report: {COLOR_THEMES[_color_idx][0]} "
                   f"(next report: {COLOR_THEMES[_next_enabled(_color_idx)][0]})")
