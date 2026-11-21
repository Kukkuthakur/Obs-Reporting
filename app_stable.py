import streamlit as st
import json
import io
import html
import math
import re
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.table import WD_ALIGN_VERTICAL
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
# ============================================================
# FIRST-TRIMESTER SCORING (CRL 45-84 mm): NT, mean uterine artery PI, ductus venosus PI
# ============================================================
from statistics import NormalDist
_ND = NormalDist()
FT_CRL_MIN, FT_CRL_MAX = 45.0, 84.0

def _ft_pct(z):
    return round(min(max(_ND.cdf(z) * 100, 0.1), 99.9), 1)

def _ft_crl(crl):
    """CRL in mm as a float, or None when it is missing / outside the 45-84 mm window the references cover."""
    try:
        c = float(str(crl).lower().replace("mm", "").strip())
    except (TypeError, ValueError):
        return None
    return c if FT_CRL_MIN <= c <= FT_CRL_MAX else None

def _ft_interp(rows, crl):
    """rows = [(crl, v1, v2, ...)] sorted by CRL -> values interpolated linearly at this CRL."""
    for a, b in zip(rows, rows[1:]):
        if a[0] <= crl <= b[0]:
            t = (crl - a[0]) / (b[0] - a[0])
            return [x + (y - x) * t for x, y in zip(a[1:], b[1:])]
    return list(rows[-1][1:])

# NT: expected (median) NT = 0.437 + 0.01969 x CRL, spread SD = 0.56 mm.
# Fitted to the readings of the Perinatology.com CRL & NT calculator (matches all but one keyed reading to +/-1 percentile).
NT_SD_MM = 0.56
def nt_expected(crl):
    return 0.437 + 0.01969 * crl

def calculate_nt_pct(nt, crl):
    """-> {'pct', 'expected'} or None when NT is missing or the CRL is outside 45-84 mm."""
    c = _ft_crl(crl)
    try:
        v = float(str(nt).lower().replace("mm", "").strip())
    except (TypeError, ValueError):
        return None
    if c is None or v <= 0:
        return None
    exp = nt_expected(c)
    return {"pct": _ft_pct((v - exp) / NT_SD_MM), "expected": exp}

# Mean uterine artery PI at 11-14 weeks by CRL: Niu et al., Quant Imaging Med Surg 2023;13:8587 (Table 2 natural, Table 3 IVF).
# Columns: CRL, P5, P10, P25, P50, P75, P90, P95 (linear between the CRLs listed).
_UTA_LEVELS = (5, 10, 25, 50, 75, 90, 95)
_UTA_REF = {
    "Natural": [(45, 1.25, 1.39, 1.62, 1.91, 2.23, 2.54, 2.74), (50, 1.22, 1.34, 1.57, 1.86, 2.16, 2.46, 2.66),
                (55, 1.18, 1.30, 1.53, 1.80, 2.10, 2.39, 2.58), (60, 1.14, 1.26, 1.48, 1.74, 2.03, 2.32, 2.50),
                (65, 1.11, 1.22, 1.43, 1.69, 1.97, 2.24, 2.42), (70, 1.07, 1.18, 1.38, 1.63, 1.90, 2.17, 2.33),
                (75, 1.03, 1.14, 1.34, 1.57, 1.84, 2.09, 2.25), (80, 1.00, 1.10, 1.29, 1.52, 1.77, 2.02, 2.17),
                (84, 0.97, 1.07, 1.25, 1.47, 1.72, 1.95, 2.11)],
    "IVF":     [(45, 1.01, 1.16, 1.43, 1.73, 2.04, 2.33, 2.50), (50, 0.98, 1.13, 1.38, 1.68, 1.98, 2.26, 2.43),
                (55, 0.95, 1.09, 1.34, 1.63, 1.92, 2.19, 2.36), (60, 0.92, 1.06, 1.30, 1.58, 1.86, 2.13, 2.29),
                (65, 0.89, 1.03, 1.26, 1.53, 1.80, 2.06, 2.21), (70, 0.86, 0.99, 1.22, 1.48, 1.74, 1.99, 2.14),
                (75, 0.83, 0.96, 1.18, 1.43, 1.69, 1.92, 2.07), (80, 0.81, 0.93, 1.14, 1.38, 1.63, 1.85, 1.99),
                (84, 0.78, 0.90, 1.11, 1.34, 1.58, 1.80, 1.94)],
}

def calculate_ft_uta_pi(pi, crl, conception="Natural"):
    """-> {'pct', 'p5', 'p95'} or None. Percentile = position between the published P5..P95 curves (z-scale), extended at the ends."""
    c = _ft_crl(crl)
    try:
        v = float(str(pi).strip())
    except (TypeError, ValueError):
        return None
    if c is None or v <= 0:
        return None
    curve = _ft_interp(_UTA_REF["IVF" if str(conception).upper() == "IVF" else "Natural"], c)
    zs = [_ND.inv_cdf(l / 100) for l in _UTA_LEVELS]
    if v <= curve[0]:
        i = 0
    elif v >= curve[-1]:
        i = len(curve) - 2
    else:
        i = next(k for k in range(len(curve) - 1) if curve[k] <= v <= curve[k + 1])
    t = (v - curve[i]) / (curve[i + 1] - curve[i])
    z = zs[i] + (zs[i + 1] - zs[i]) * t
    return {"pct": _ft_pct(z), "p5": curve[0], "p95": curve[-1]}

# Ductus venosus PI at 11-14 weeks by CRL: Li et al., PLoS One 2024;19:e0312874 (Table 1): P5 and P95 per CRL.
# Mean = (P5+P95)/2 and SD = (P95-P5)/(2 x 1.6449), as the paper derives them.
_DV_REF = [(45, 1.0007, 1.3415), (50, 0.9790, 1.3349), (55, 0.9617, 1.3258), (60, 0.9496, 1.3140), (65, 0.9432, 1.2993),
           (70, 0.9430, 1.2815), (75, 0.9488, 1.2602), (81, 0.9633, 1.2294), (84, 0.9734, 1.2115)]

def calculate_ft_dv_pi(pi, crl):
    """-> {'pct', 'p5', 'p95'} or None."""
    c = _ft_crl(crl)
    try:
        v = float(str(pi).strip())
    except (TypeError, ValueError):
        return None
    if c is None or v <= 0:
        return None
    p5, p95 = _ft_interp(_DV_REF, c)
    mean, sd = (p5 + p95) / 2, (p95 - p5) / (2 * 1.6449)
    return {"pct": _ft_pct((v - mean) / sd), "p5": p5, "p95": p95}

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
    'crown rump length': ('crl', False),   # -> crl_wks / crl_dys / crl_pct (GA and percentile of the CRL)
    'nasal bone': ('nb', True),            # -> nb_pct (the machine's own nasal-bone percentile)
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

# --- TWINS: split an SR by "Fetus ID" so each fetus keeps its own measurements ---
# Fetus 1 -> Twin A, Fetus 2 -> Twin B. Containers without a Fetus ID (uterine arteries, dates, patient) belong to both.
def _sr_child_text(item, meaning):
    """Value of the TEXT child whose concept name is `meaning` (e.g. 'fetus id'), else None."""
    for sub in (item.get('ContentSequence', []) or []):
        if 'ConceptNameCodeSequence' in sub and \
           _sr_norm(sub.ConceptNameCodeSequence[0].get('CodeMeaning', '')) == meaning and sub.get('ValueType') == 'TEXT':
            return str(sub.get('TextValue', '')).strip()
    return None

def _sr_fetus_ids(ds):
    ids = []
    def walk(seq):
        for it in seq:
            fid = _sr_child_text(it, 'fetus id')
            if fid and fid not in ids:
                ids.append(fid)
            if 'ContentSequence' in it:
                walk(it.ContentSequence)
    if 'ContentSequence' in ds:
        walk(ds.ContentSequence)
    return sorted(ids, key=lambda x: (len(x), x))

def _sr_number_of_fetuses(ds):
    found = []
    def walk(seq):
        for it in seq:
            if 'ConceptNameCodeSequence' in it and it.get('ValueType') == 'NUM' and 'MeasuredValueSequence' in it \
               and _sr_norm(it.ConceptNameCodeSequence[0].get('CodeMeaning', '')) == 'number of fetuses':
                try:
                    found.append(int(float(it.MeasuredValueSequence[0].get('NumericValue', 0))))
                except (TypeError, ValueError):
                    pass
            if 'ContentSequence' in it:
                walk(it.ContentSequence)
    if 'ContentSequence' in ds:
        walk(ds.ContentSequence)
    return max(found) if found else 0

def _sr_only_fetus(ds, fid):
    """Copy of the SR with every container that belongs to a different fetus removed."""
    import copy
    from pydicom.sequence import Sequence
    ds2 = copy.deepcopy(ds)
    def prune(seq):
        keep = []
        for it in seq:
            own = _sr_child_text(it, 'fetus id')
            if own and own != fid:
                continue
            if 'ContentSequence' in it:
                it.ContentSequence = prune(it.ContentSequence)
            keep.append(it)
        return Sequence(keep)
    if 'ContentSequence' in ds2:
        ds2.ContentSequence = prune(ds2.ContentSequence)
    return ds2

def extract_sr_data(ds):
    """One fetus: the usual flat dict (+ n_fetuses, fetus_id).
    Twins in one SR: flat dict of the first fetus, plus 'fetuses': {'1': {...}, '2': {...}} and n_fetuses."""
    ids = _sr_fetus_ids(ds)
    n = _sr_number_of_fetuses(ds)
    if len(ids) < 2:
        out = _extract_sr_flat(ds)
        out['n_fetuses'] = max(n, 1)
        if ids:
            out['fetus_id'] = ids[0]
        return out
    per = {fid: _extract_sr_flat(_sr_only_fetus(ds, fid)) for fid in ids}
    out = dict(per[ids[0]])
    out['fetuses'] = per
    out['n_fetuses'] = max(n, len(ids))
    out['fetus_id'] = ids[0]
    return out

def _extract_sr_flat(ds):
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

        # --- DATE VALUES (LMP and EDDs are DATE items on this machine) ---
        if item.ValueType == 'DATE':
            _raw = str(item.get('Date', '') or '').strip()
            if re.fullmatch(r'\d{8}', _raw):
                _dd = f"{_raw[6:8]}-{_raw[4:6]}-{_raw[0:4]}"
                if concept == 'EDD from average ultrasound age': extracted['edd_cua'] = _dd
                elif concept == 'EDD from LMP': extracted['edd_lmp'] = _dd
                elif concept == 'LMP': extracted['lmp'] = _dd

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
                elif concept == 'Crown Rump Length': extracted['crl'] = f"{int(round(num_val * 10))}"

            # Anomaly-scan sizes: TCD, lateral ventricle, cisterna magna, nuchal fold, radius, ulna, tibia, fibula
            if _bk and _bk[0] in SR_EXTRA_SIZE_KEYS:
                _mm = _sr_size_mm(num_val, unit, _bk[1])
                if _mm is not None:
                    extracted[_bk[0]] = _mm

            # NT arrives in mm on some machines and cm on others; keep two decimals (1.27 mm)
            if _sr_norm(concept) == 'nuchal translucency':
                _ntmm = num_val * 10 if unit == 'centimeter' else num_val if unit == 'millimeter' else None
                if _ntmm is not None:
                    extracted['nt'] = f"{_ntmm:.2f}".rstrip('0').rstrip('.')
            # Nasal bone: length (mm) and Present/Absent; the percentile comes from the sibling 'Growth Percentile Rank'
            if _sr_norm(concept) == 'nasal bone':
                extracted['nasal_bone'] = "Present" if num_val > 0 else "Absent"
                _nbmm = _sr_size_mm(num_val, unit, True)
                if _nbmm is not None:
                    extracted['nb_len'] = _nbmm
            
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

    # Date calculations: counted from the scan date in the SR (falls back to today)
    today = datetime.now()
    _sd = str(getattr(ds, 'StudyDate', '') or '').strip()
    if re.fullmatch(r'\d{8}', _sd):
        try:
            today = datetime.strptime(_sd, "%Y%m%d")
            extracted['study_date'] = today.strftime("%d-%b-%y")
        except ValueError:
            pass
    _rp = str(getattr(ds, 'ReferringPhysicianName', '') or '').replace('^', ' ').strip()
    if _rp:
        extracted['referred_by'] = _rp
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

def align_tables_left(doc):
    """Every table starts at exactly the same left edge (the page's left margin): left-aligned, zero indent, and Word's
    modern table-position mode so the border (not the text) sits on the margin. Table widths are NOT touched."""
    st_el = doc.settings.element
    compat = st_el.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat'); st_el.append(compat)
    for cs in compat.findall(qn('w:compatSetting')):
        if cs.get(qn('w:name')) == 'compatibilityMode': compat.remove(cs)
    cs = OxmlElement('w:compatSetting')
    cs.set(qn('w:name'), 'compatibilityMode'); cs.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word'); cs.set(qn('w:val'), '15')
    compat.append(cs)
    for tbl in doc.tables:
        pr = tbl._tbl.tblPr
        for tag in ('w:jc', 'w:tblInd'):
            for e in pr.findall(qn(tag)): pr.remove(e)
        j = OxmlElement('w:jc'); j.set(qn('w:val'), 'left'); pr.append(j)
        i = OxmlElement('w:tblInd'); i.set(qn('w:type'), 'dxa'); i.set(qn('w:w'), '0'); pr.append(i)

def keep_row_together(table):
    """Do not let a table row split across two pages (the PNDT box must never be cut in half)."""
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

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
TEMPLATE_ALLOWED_THEMES = {"TWINS_EARLY": {"Pink", "Lavender"},   # twin Early OBS rotates between these two only
                           "EARLY_OBS": {"Pink", "Lavender"},     # singleton Early OBS: Pink / Lavender only
                           "ANOMALY": {"Blue", "Pink"}}           # singleton Anomaly Scan: Blue / Pink only (never Lavender)
TEMPLATE_DISABLED_THEMES = {"NT_NB": {"Blue"}}   # colours a particular template never uses (the rotation is otherwise unchanged)
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

def theme_for_template(template_id, idx):
    """Colour a report of this template actually uses: the rotation colour, or the next allowed one if the
    template does not use it (e.g. NT/NB never uses Blue)."""
    blocked = TEMPLATE_DISABLED_THEMES.get(template_id, set())
    allowed = TEMPLATE_ALLOWED_THEMES.get(template_id)
    for step in range(len(COLOR_THEMES) + 1):
        j = (idx + step) % len(COLOR_THEMES)
        if j in ROTATION and COLOR_THEMES[j][0] not in blocked and (allowed is None or COLOR_THEMES[j][0] in allowed):
            return j
    return idx

def advance_color_index(used_idx=None):
    """Called when a report is downloaded: the next report gets the colour after the one just used."""
    import json
    try:
        next_index = _next_enabled(load_color_index() if used_idx is None else used_idx)   # read BEFORE the file is reopened for writing
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

def oligo_impression_phrase(data):
    """Impression fragment for abnormal liquor (oligo / nil / polyhydramnios), else ''.
    Total 4 Quadrants: follows generate_afi_text bands (0 nil; <4 moderate to severe; 4-<6 mild to moderate;
    6-<8 mild oligohydramnios; >=25 polyhydramnios).
    Largest Pocket: <2 cm oligohydramnios; >8 cm polyhydramnios."""
    method = data.get('afi_method', 'Total 4 Quadrants')
    if method == "Total 4 Quadrants":
        val = str(data.get('afi_total', '') or '').strip()
        try:
            f = float(val)
        except ValueError:
            return ""
        desc, _ = generate_afi_text(val, is_total=True)
        if f == 0:
            return "anhydramnios, not a single pocket of even one cm seen"
        if 0 < f < 8 and "oligohydramnios" in desc:
            return f"{desc.lower()}, total of four quadrants measures {val} cm"
        if desc == "Polyhydramnios":
            return f"polyhydramnios, total of four quadrants measures {val} cm"
        return ""
    val = str(data.get('afi_pocket', '') or '').strip()
    try:
        f = float(val)
    except ValueError:
        return ""
    if f == 0:
        return "anhydramnios, not a single pocket of even one cm seen"
    if 0 < f < 2.0:
        return f"oligohydramnios, largest pocket measures {val} cm"
    if f > 8.0:
        return f"polyhydramnios, largest pocket measures {val} cm"
    return ""

def impression_with_clause(inline, data):
    """' with ...' clause for the impression line: oligohydramnios first, then the other findings."""
    oligo = oligo_impression_phrase(data)
    inline = [i for i in inline if i]
    if oligo and inline:
        return " with " + oligo + ", and " + _join_and(inline)
    if oligo:
        return " with " + oligo
    if inline:
        return " with " + _join_and(inline)
    return ""

# ============================================================
# DOCX BUILDER
# ============================================================
PLACENTA_OPTS = ["Anterior", "Posterior", "Fundo-Anterior", "Fundo-Posterior",
                 "Fundal extending antero-posteriorly"]

# Low-lying placenta: how it relates to the Internal Os. The chosen text is appended as it is after "low-lying" in the
# placenta line; options 2-4 also add an advisory under the impression. Option 1 is the default.
PLAC_LL_OPTS = [", but away from the Internal Os.", ", reaching upto the Internal Os.",
                ", partially covering the Internal Os.", ", completely covering the Internal Os."]
PLAC_LL_ADVICE = "Adv- Review scan after a month to assess placental migration."

def lowlying_text(d):
    """Text that follows 'low-lying'. Default = option 1. (A distance typed on the standard form keeps option 1 as 'about X cm away'.)"""
    ch = d.get("placenta_ll") or PLAC_LL_OPTS[0]
    if ch not in PLAC_LL_OPTS:
        ch = PLAC_LL_OPTS[0]
    dist = str(d.get("placenta_dist", "") or "").strip()
    if ch == PLAC_LL_OPTS[0] and dist:
        return f", about {dist}cm away from the Internal Os."
    return ch

def lowlying_advice(d):
    """Advisory line for a low-lying placenta that reaches / partly covers / completely covers the Internal Os (options 2-4)."""
    if d.get("placenta_os") == "Low-lying" and lowlying_text(d) in PLAC_LL_OPTS[1:]:
        return PLAC_LL_ADVICE
    return ""

def lowlying_select(container, key, label="Low-lying placenta - relation to Internal Os"):
    return container.selectbox(label, PLAC_LL_OPTS, index=0, key=key,
                               format_func=lambda o: f"{PLAC_LL_OPTS.index(o) + 1}. {o}")

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

def _mm_text(v):
    s = str(v or "").strip().lower().replace("mm", "").strip()
    return f"{s} mm" if s else ""

def cpc_phrase(data):
    """Choroid plexus cyst wording, e.g. 'a right choroid plexus cyst (5 mm)' / 'Bilateral choroid plexi cysts (5 mm & 4 mm)'."""
    if not data.get("cpc"):
        return ""
    r, l = bool(data.get("cpc_right")), bool(data.get("cpc_left"))
    rs, ls = _mm_text(data.get("cpc_right_size")), _mm_text(data.get("cpc_left_size"))
    if r and l:
        sizes = " & ".join(x for x in (rs, ls) if x)          # right first, then left
        return "Bilateral choroid plexi cysts" + (f" ({sizes})" if sizes else "")
    if r or l:
        size = rs if r else ls
        return f"a {'right' if r else 'left'} choroid plexus cyst" + (f" ({size})" if size else "")
    return ""

def eif_phrase(data):
    """Echogenic intracardiac focus wording."""
    if not data.get("eif"):
        return ""
    r, l = bool(data.get("eif_right")), bool(data.get("eif_left"))
    if r and l:
        return "Echogenic intracardiac foci (EIF) in both Left & Right ventricles"
    if r or l:
        return f"an Echogenic intracardiac focus (EIF) in {'Right' if r else 'Left'} ventricle"
    return ""

def softmarker_notes(data):
    """Report-body lines. Anomaly: replace 'No identifiable intracranial lesion seen.' / 'No evidence of SOL in the thorax.'
    T16 / T3 / Color Doppler: extra last lines under the cord / scar line."""
    return [f"Note is made of {ph}." for ph in (cpc_phrase(data), eif_phrase(data)) if ph]

def softmarker_phrases(data):
    return [ph for ph in (cpc_phrase(data), eif_phrase(data)) if ph]

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
    m = re.search(r'(\d+)\s*w(?:eeks?|ks?)?\s*(?:(\d+)\s*d)?', str(s or '').lower())
    return (int(m.group(1)), int(m.group(2) or 0)) if m else None

def sr_ga_days(sr):
    """Gestational age in days from the DICOM SR: composite ultrasound age (mean age) first, then ga_usg_days,
    then GA by LMP. Twin SR files may nest per-fetus dicts, so nested dicts are searched too. None if absent."""
    def _find(d, depth=0):
        if not isinstance(d, dict) or depth > 3:
            return None
        ga = _split_ga(d.get("mean_age")) if d.get("mean_age") else None
        if ga:
            return ga[0] * 7 + ga[1]
        for k in ("ga_usg_days", "ga_lmp_days"):
            try:
                if d.get(k) not in (None, ""):
                    return int(float(d[k]))
            except (TypeError, ValueError):
                pass
        for val in d.values():
            if isinstance(val, dict):
                r = _find(val, depth + 1)
                if r is not None:
                    return r
        return None
    return _find(sr or {})

def auto_obs_template(sr, twins=False):
    """GA below 25w0d -> T16, 25w0d or more -> T3. None when the SR has no gestational age."""
    days = sr_ga_days(sr)
    if days is None:
        return None
    return ("TWINS_" if twins else "") + ("T16" if days < 25 * 7 else "T3")


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

def nt_nb_summary(data):
    """Everything the NT/NB report prints, worked out once: texts, percentiles and the abnormal findings for the opinion."""
    S = {"flags": []}
    ga = _split_ga(data.get('ga_crl')) or _split_ga(data.get('mean_age'))
    S["ga_txt"] = f"{ga[0]}wks {ga[1]:02d}days" if ga else ""
    crl = str(data.get('crl', '') or '').strip()
    S["crl_line"] = (f"{crl} mm" + (f" - {S['ga_txt']}" if S["ga_txt"] else "")) if crl else S["ga_txt"]
    fhr = str(data.get('fhr_val', '') or '').strip()
    if data.get('cardiac_status', 'Present') == "Present":
        S["cardiac"] = "Present" + (f", {fhr} bpm" if fhr else "")
    else:
        S["cardiac"] = "Not appreciated"
    S["subchorionic"] = "Absent" if not data.get("subchorionic_present") else f"Present; {data.get('subchorionic_grade', 'Grade-I')}"
    S["corpus"] = ("Not appreciated" if not data.get("corpus_present")
                   else f"Appreciated; {data.get('corpus_size', '')}mm in {data.get('corpus_side', 'Right')} Ovary")
    S["cervix"] = fmt_cm(data.get('cervix', '')) or "Normal"
    S["internal_os"] = data.get("internal_os", "Closed")
    _pl = str(data.get("placenta_pos", "") or "").strip()
    _pl_map = {"Anterior": "Anteriorly", "Posterior": "Posteriorly", "Fundo-Anterior": "Fundo-Anteriorly",
               "Fundo-Posterior": "Fundo-Posteriorly"}
    S["placenta"] = ("Forming " + _pl_map.get(_pl, _pl)) if _pl else "-"
    if data.get("placenta_os") == "Low-lying" and S["placenta"] != "-":
        S["placenta"] += ", low-lying" + lowlying_text(data)
    elif data.get("placenta_os") and data.get("placenta_os") != "away from the Os":
        S["placenta"] += f", {data.get('placenta_os')}"
    S["choroid"] = data.get("choroid", "Appreciated, appear Normal")

    # --- NT ---
    nt = str(data.get('nt', '') or '').strip()
    S["nt_pct"], S["nt_pct_note"], S["nt_val"] = None, "", "-"
    if nt:
        try:
            ntf = float(nt)
        except ValueError:
            ntf = None
        res = calculate_nt_pct(nt, crl)
        inc = bool((res and res["pct"] >= 95) or (ntf and ntf >= 3.5))
        status = "Increased" if inc else ("WNL" if res else "")
        S["nt_val"] = f"{nt}mm" + (f" ({status})" if status else "")
        if res:
            S["nt_pct"] = res["pct"]
        else:
            S["nt_pct_note"] = "n/a (CRL outside 45-84 mm)"
        if inc:
            S["flags"].append(f"increased nuchal translucency ({nt}mm)")

    # --- nasal bone (length and percentile come from the SR / machine) ---
    nbs = str(data.get('nasal_bone', '') or '').strip()
    nbl = str(data.get('nb_len', '') or '').strip()
    nbp = str(data.get('nb_pct', '') or '').strip()
    S["nb_pct"] = None
    try:
        S["nb_pct"] = float(nbp) if nbp else None
    except ValueError:
        pass
    if nbs.lower() == "absent":
        S["nb_val"] = "Absent"
        S["flags"].append("absent nasal bone")
    elif nbl:
        short = S["nb_pct"] is not None and S["nb_pct"] < 5
        S["nb_val"] = f"{nbl}mm" + (" (Short)" if short else " (WNL)" if S["nb_pct"] is not None else "")
        if short:
            S["flags"].append(f"short nasal bone ({nbl}mm)")
    else:
        S["nb_val"] = nbs or "-"

    # --- pre-eclampsia screen: right and left uterine arteries, then the mean ---
    S["uta"] = None

    def _uf(k):
        return str(data.get(k, '') or '').strip()

    def _avg(*vals):
        nums = []
        for x in vals:
            try:
                nums.append(float(x))
            except (TypeError, ValueError):
                pass
        return f"{sum(nums) / len(nums):.2f}" if nums else ""

    def _n2(x):
        try:
            return f"{float(x):.2f}"
        except (TypeError, ValueError):
            return x or "-"

    uta_rows = []
    for lab, side in (("Rt. Uterine Artery", "uterine_right"), ("Lt. Uterine Artery", "uterine_left")):
        if any(_uf(f"{side}_{m}") for m in ("pi", "ri", "sd")):
            uta_rows.append((lab, _n2(_uf(f"{side}_pi")), _n2(_uf(f"{side}_ri")), _n2(_uf(f"{side}_sd"))))
    pi = _uf('uterine_pi') or _avg(_uf('uterine_left_pi'), _uf('uterine_right_pi'))
    ri = _uf('uterine_ri') or _avg(_uf('uterine_left_ri'), _uf('uterine_right_ri'))
    sdv = _uf('uterine_sd') or _avg(_uf('uterine_left_sd'), _uf('uterine_right_sd'))
    notch = str(data.get('uterine_notching', 'No notching') or 'No notching')
    if pi or uta_rows:
        res = calculate_ft_uta_pi(pi, crl, data.get('conception', 'Natural')) if pi else None
        raised = bool(res and res["pct"] > 95)
        notched = notch.lower() != "no notching"
        text = (f"{'Raised impedance' if raised else 'Normal flow characteristics'} seen in the uterine arteries, "
                f"{notch.lower()} seen.")
        S["uta"] = {"text": text, "status": "Raised" if raised else ("WNL" if res else ""),
                    "pct": res["pct"] if res else None,
                    "note": "" if (res or not pi) else "percentile n/a - CRL outside 45-84 mm",
                    "rows": uta_rows, "mean": (_n2(pi) if pi else "", _n2(ri) if ri else "", _n2(sdv) if sdv else "")}
        if raised:
            S["flags"].append(f"raised mean uterine artery PI ({pi})")
        if notched:
            S["flags"].append(f"{notch.lower()} in the uterine arteries")

    # --- ductus venosus ---
    S["dv"] = None
    dvp = str(data.get('dv_pi', '') or '').strip()
    wave = data.get('dv_awave', 'Positive')
    if dvp:
        res = calculate_ft_dv_pi(dvp, crl)
        raised = bool(res and res["pct"] > 95)
        bad_wave = wave in ("Absent", "Reversed")
        wave_txt = "Normal S,D & A waves" if not bad_wave else f"{wave} a-wave"
        text = f"{'Abnormal' if (raised or bad_wave) else 'Normal'} flow characteristics seen, PI- {dvp} ({wave_txt})"
        S["dv"] = {"text": text, "status": "Raised" if raised else "", "pct": res["pct"] if res else None,
                   "note": "" if res else "percentile n/a - CRL outside 45-84 mm"}
        if raised:
            S["flags"].append(f"raised ductus venosus PI ({dvp})")
        if bad_wave:
            S["flags"].append(f"{wave.lower()} a-wave in the ductus venosus")
    return S

def nt_nb_opinion(data):
    S = nt_nb_summary(data)
    ga = S["ga_txt"]
    if data.get('cardiac_status', 'Present') == "Present":
        txt = f"Intra uterine live pregnancy of {ga} +/-1w with good cardiac activity"
    else:
        txt = f"Intra uterine pregnancy of {ga} +/-1w with no cardiac activity appreciated"
    if S["flags"]:
        txt += ", with " + _join_and(S["flags"])
    return txt + "."

def nt_nb_preview(data):
    S = nt_nb_summary(data)
    L = [f"LMP- {data.get('lmp', '')}", "",
         f"  {'CRL':<26}: {S['crl_line']}", f"  {'Cardiac pulsations':<26}: {S['cardiac']}",
         f"  {'Subchorionic collection':<26}: {S['subchorionic']}", f"  {'Corpus luteal cyst':<26}: {S['corpus']}",
         f"  {'Cervix':<26}: {S['cervix']}", f"  {'Internal Os':<26}: {S['internal_os']}",
         f"  {'Placenta':<26}: {S['placenta']}", "", "FETAL CRANIUM:",
         f"  {'Choroid plexi / Falx':<26}: {S['choroid']}",
         f"  {'Nuchal translucency (NT)':<26}: {S['nt_val']}"
         + (f"  [{S['nt_pct']:g}%]" if S['nt_pct'] is not None else (f"  [{S['nt_pct_note']}]" if S['nt_pct_note'] else "")),
         f"  {'Nasal bone (NB)':<26}: {S['nb_val']}" + (f"  [{S['nb_pct']:g}%]" if S['nb_pct'] is not None else "")]
    def _line(head, d):
        if not d:
            return []
        extra = (" (" + ", ".join([x for x in (d["status"], f"{d['pct']:g}% for CRL") if x]) + ")") if d["pct"] is not None \
            else (f" ({d['note']})" if d["note"] else "")
        return ["", head, f"  {d['text']}{extra}."]
    if S["uta"]:
        d = S["uta"]
        L += ["", "PRE-ECLAMPSIA SCREEN:", f"  {d['text']}", f"  {'Uterine artery':<24}{'PI':>7}{'RI':>7}{'S/D':>7}"]
        for r in d["rows"]:
            L.append(f"  {r[0]:<24}{r[1]:>7}{r[2]:>7}{r[3]:>7}")
        m = d["mean"]
        tail = (f"   ({d['status']}, {d['pct']:g}% for CRL)" if d["pct"] is not None
                else (f"   ({d['note']})" if d["note"] else ""))
        L.append(f"  {'Mean Uterine Arteries':<24}{m[0] or '-':>7}{m[1] or '-':>7}{m[2] or '-':>7}{tail}")
    L += _line("DUCTUS VENOSUS:", S["dv"])
    L += [""]
    for lab, key in (("EDD (LMP)", "edd_lmp"), ("EDD (CUA)", "edd_cua")):
        if data.get(key):
            L.append(f"  {lab:<10}: {_fmt_date(data.get(key))}  +/-2W")
    return L

# ============================================================
# EARLY OBS (singleton) - same look as NT/NB (Pink / Lavender), own findings rows, no cranium / DV
# ============================================================
EARLY_CL_LABEL = "Corpus leuteal cyst"        # spelled exactly as in the approved sample report

def _num_only(v):
    """'13', '13mm', ' 13.5 mm ' -> '13' / '13.5'; '' if there is no number."""
    m = re.search(r'\d+(?:\.\d+)?', str(v or ""))
    return m.group(0) if m else ""

def parse_cl_cyst(text):
    """Machine annotation read by the OCR -> (side, size_mm).
    'CL CYST RT OVARY' -> ('Right', '');  'CL CYST LT OVARY' -> ('Left', '');  both annotations -> ('Both', '').
    A size written in the annotation (e.g. '2.1 cm' / '21 mm') is returned in mm. (None, '') if no side is found."""
    t = str(text or "").upper()
    right = bool(re.search(r'\b(RT|RIGHT)\b', t))
    left = bool(re.search(r'\b(LT|LEFT)\b', t))
    side = "Both" if (right and left) else "Right" if right else "Left" if left else None
    size = ""
    m = re.search(r'(\d+(?:\.\d+)?)\s*(MM|CM)\b', t)
    if m and side in ("Right", "Left"):
        v = float(m.group(1)) * (10 if m.group(2) == "CM" else 1)
        size = f"{v:g}"
    return side, size

def early_corpus_text(data):
    if not data.get("corpus_present"):
        return "Not appreciated"
    side = data.get("corpus_side", "Right")
    size = _num_only(data.get("corpus_size"))
    if side == "Both":
        return "Appreciated; in both Ovaries"
    return f"Appreciated; {size}mm in {side} Ovary" if size else f"Appreciated; in {side} Ovary"

def early_cervix_text(val):
    """'4.4' / '44 mm' -> 'Normal(4.4cm).'; below 2.5 cm -> 'Short(2.1cm).'; typed text is kept as typed."""
    s = fmt_cm(val)
    if not s:
        return "Normal."
    m = re.fullmatch(r'(\d+(?:\.\d+)?) cm', s)
    if not m:
        return "Normal." if s.strip().lower() == "normal" else s
    return f"{'Normal' if float(m.group(1)) >= 2.5 else 'Short'}({m.group(1)}cm)."

def early_summary(data):
    """nt_nb_summary() (placenta, subchorionic, uterine arteries, GA, flags) plus the Early OBS findings rows."""
    S = nt_nb_summary(data)
    msd = _num_only(data.get("msd"))
    # G-Sac: only written as 'Seen' unless the sac was measured (GS / MSD) - then the measurement appears
    S["gsac"] = f"Seen, MSD {msd}mm" if msd else "Seen"
    crl = _num_only(data.get("crl"))
    S["crl_txt"] = (f"{crl}mm- {S['ga_txt']}" if S["ga_txt"] else f"{crl}mm") if crl else "Not seen"
    status = data.get("cardiac_status", "Present")
    fhr = _num_only(data.get("fhr_val"))
    S["cardiac_txt"] = (f"Present {fhr} bpm" if fhr else "Present") if status == "Present" \
        else "Not yet appreciable" if status == "Not yet appreciable" else "Absent"
    S["yolk_txt"] = str(data.get("yolk", "") or "Not seen").strip() or "Not seen"
    S["decidual_txt"] = data.get("decidual", "Present")
    S["corpus_txt"] = early_corpus_text(data)
    S["cervix_txt"] = early_cervix_text(data.get("cervix", ""))
    S["early_rows"] = [("G-Sac", S["gsac"]), ("CRL", S["crl_txt"]), ("Yolk sac", S["yolk_txt"]),
                       ("Cardiac pulsations", S["cardiac_txt"]), ("Decidual reaction", S["decidual_txt"]),
                       ("Subchorionic collection", S["subchorionic"]), (EARLY_CL_LABEL, S["corpus_txt"]),
                       ("Cervix", S["cervix_txt"]), ("Internal Os", S["internal_os"]), ("Placenta", S["placenta"])]
    return S

def early_opinion(data):
    S = nt_nb_summary(data)
    ga, status = S["ga_txt"], data.get("cardiac_status", "Present")
    extra = (", with " + _join_and(S["flags"])) if S["flags"] else ""
    if status == "Blighted Ovum":
        return ("INTRA UTERINE EMPTY GESTATIONAL SAC" + (f" OF {ga} +/-1w" if ga else "")
                + " WITH NO APPRECIABLE FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY FAILURE (BLIGHTED OVUM).")
    if status == "Not yet appreciable":
        return (f"Intra uterine pregnancy of {ga} +/-1w with not yet appreciable cardiac activity{extra}."
                " Adv- Review scan after 2 weeks.")
    return f"Intra uterine live pregnancy of {ga} +/-1w with good cardiac activity{extra}."

def early_preview(data):
    S = early_summary(data)
    L = [f"LMP- {data.get('lmp', '')}", ""]
    L += [f"  {lab:<28}: {val}" for lab, val in S["early_rows"]]
    if S["uta"]:
        d = S["uta"]
        L += ["", "PRE-ECLAMPSIA SCREEN:", f"  {d['text']}", f"  {'Uterine artery':<24}{'PI':>7}{'RI':>7}{'S/D':>7}"]
        for r in d["rows"]:
            L.append(f"  {r[0]:<24}{r[1]:>7}{r[2]:>7}{r[3]:>7}")
        m = d["mean"]
        tail = (f"   ({d['status']}, {d['pct']:g}% for CRL)" if d["pct"] is not None
                else (f"   ({d['note']})" if d["note"] else ""))
        L.append(f"  {'Mean Uterine Arteries':<24}{m[0] or '-':>7}{m[1] or '-':>7}{m[2] or '-':>7}{tail}")
    L.append("")
    for lab, key in (("EDD (LMP)", "edd_lmp"), ("EDD (CUA)", "edd_cua")):
        if data.get(key) and not (key == "edd_lmp" and str(data.get("lmp", "")).strip() == "Not Sure"):
            L.append(f"  {lab:<10}: {_fmt_date(data.get(key))}  +/-2W")
    return L

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
    _gap_pt(doc, 4)

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
    _gap_pt(doc, 6)

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
    p.paragraph_format.space_before = Pt(4)
    add_run(p, "Uterine artery notching: ", bold=True)
    add_run(p, str(data.get("uterine_notching", "No notching")), italic=True)

    p = doc.add_paragraph()
    add_run(p, DOPPLER_REF_PREFIX, italic=True, size=8)
    add_run(p, "Interpretation key: ", bold=True, size=8)
    add_run(p, "umbilical and uterine artery PI above the 95th percentile, and MCA PI or cerebro-placental ratio "
               "below the 5th percentile, are considered abnormal."
               + (" MCA PSV of 1.5 MoM or more is also considered abnormal (risk of fetal anaemia)."
                  if data.get("mca_psv") else ""), italic=True, size=8)
    if data.get("cpr_pct") and not CPR_PERCENTILE_VALIDATED:
        p = doc.add_paragraph()
        add_run(p, "\u2020 CPR percentile is provisional.", italic=True, size=8)

# Small-print line under the Doppler tables starts with this reference note (singleton and twin Color Doppler).
DOPPLER_REF_PREFIX = "\u0166 Percentile reference source: Barcelona Obs Calculator. "

def _gap_pt(doc, pts):
    """Blank line of an exact height (a plain empty paragraph is a full 11 pt line)."""
    g = doc.add_paragraph()
    g.paragraph_format.space_before = g.paragraph_format.space_after = Pt(0)
    g.paragraph_format.line_spacing = Pt(pts)
    return g

def _add_page_two_break(doc, tight=False):
    """'P.T.O' marker, then start page 2 with the same top spacer as page 1.
    tight=True (Doppler pages, which are full): less space above P.T.O, so it can never be pushed alone onto the next page."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(3 if tight else 10)
    p.paragraph_format.keep_together = True
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

# Top heading of each template (printed in the navy title bar above the patient details box)
REPORT_TITLES = {
    "EARLY_OBS":     "USG OBS (EARLY PREGNANCY SCAN)",
    "NT_NB":         "USG OBS (NT / NB SCAN)",
    "T16":           "USG OBS (SECOND TRIMESTER SCAN)",
    "T3":            "USG OBS (THIRD TRIMESTER SCAN)",
    "COLOR_DOPPLER": "USG OBS (COLOR DOPPLER)",
}

def add_title_demo_box(doc, data, title, width_cm=17.0, fill=None):
    """Navy title bar on top, patient details underneath in one bordered box (matches the Anomaly Scan report)."""
    C = WD_ALIGN_PARAGRAPH.CENTER
    half = width_cm / 2
    t = _an_table(doc, 3, [half, half])
    bar = t.cell(0, 0).merge(t.cell(0, 1))
    set_cell_background(bar, fill or AN_NAVY)
    bar.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    bp = bar.paragraphs[0]
    bp.alignment = C
    add_run(bp, title, bold=True, size=12, color=RGBColor(0x1F, 0x4E, 0x79))
    age = str(data.get("patient_age", "") or "").strip()
    agesex = f"{age}Y/{data.get('p_sex', 'F')}" if age else ""
    cells = {(1, 0): f"NAME :-  {data.get('patient_name', '')}", (1, 1): f"DATE :-  {data.get('report_date', '')}",
             (2, 0): f"AGE/SEX :-  {agesex}", (2, 1): f"REF. BY :-  {data.get('referred_by', '')}"}
    for (r, c), txt in cells.items():
        t.cell(r, c).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        add_run(t.cell(r, c).paragraphs[0], txt, bold=True)
    return t

# ============================================================
# IUD (INTRAUTERINE FETAL DEMISE) TEMPLATE  - plain layout, same for 2nd and 3rd trimester
# ============================================================
IUD_OBS = "IUD_OBS"
PRESENTATION_DEFAULTS[IUD_OBS] = "CEPHALIC"
IUD_TEAL = RGBColor(0x31, 0x84, 0x9B)
# (key, checkbox label, line in the body of the report, phrase in the impression)
IUD_FINDINGS = [
    ("ebl",        "Echogenic bowel loops",
     "ECHOGENIC BOWEL LOOPS NOTED.", "ECHOGENIC BOWEL LOOPS"),
    ("spalding",   "Positive Spalding sign (overlapping cranial bones)",
     "POSITIVE SPALDING SIGN (OVERLAPPING OF CRANIAL BONES) NOTED.", "POSITIVE SPALDING SIGN"),
    ("ascites",    "Fetal ascites",
     "FETAL ASCITES NOTED.", "FETAL ASCITES"),
    ("pleural",    "Fetal pleural effusion",
     "FETAL PLEURAL EFFUSION NOTED.", "FETAL PLEURAL EFFUSION"),
    ("scalp",      "Scalp oedema (halo sign)",
     "SCALP OEDEMA (HALO SIGN) NOTED.", "SCALP OEDEMA (HALO SIGN)"),
    ("skin",       "Generalised skin / subcutaneous oedema",
     "GENERALISED SKIN / SUBCUTANEOUS OEDEMA NOTED.", "GENERALISED SKIN OEDEMA"),
    ("gas",        "Gas in fetal heart / great vessels (Robert's sign)",
     "GAS WITHIN THE FETAL HEART / GREAT VESSELS (ROBERT'S SIGN) NOTED.", "GAS WITHIN FETAL HEART (ROBERT'S SIGN)"),
    ("spine",      "Hyperflexion / abnormal angulation of fetal spine",
     "HYPERFLEXION / ABNORMAL ANGULATION OF FETAL SPINE NOTED.", "ABNORMAL ANGULATION OF FETAL SPINE"),
    ("maceration", "Loss of fetal anatomical details (maceration)",
     "LOSS OF NORMAL FETAL ANATOMICAL DETAILS (MACERATION) NOTED.", "MACERATION OF THE FETUS"),
]

def iud_findings(data):
    """Ticked findings (plus the free-text one) as [(body line, impression phrase)], in list order."""
    out = [(body, imp) for key, _lbl, body, imp in IUD_FINDINGS if data.get(f"iud_{key}")]
    other = str(data.get("iud_other", "") or "").strip().rstrip(".").strip()
    if other:
        out.append((other.upper() + " NOTED." if not other.upper().endswith("NOTED") else other.upper() + ".",
                    other.upper().replace(" NOTED", "")))
    return out

def _iud_age(data, upper=True):
    ga = _split_ga(data.get("mean_age"))
    if ga:
        return f"{ga[0]}WEEKS {ga[1]:02d}DAYS \u00b1 2W" if upper else f"{ga[0]}weeks {ga[1]:02d}days \u00b1 2w"
    raw = str(data.get("mean_age", "") or "")
    return raw.upper() if upper else raw

def _iud_lmp(data):
    lmp = str(data.get("lmp", "") or "").strip()
    return "NOT KNOWN" if (not lmp or lmp.lower() in ("not sure", "not known")) else lmp

def _iud_join(items):
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " & " + items[-1]

def iud_impression_parts(data):
    """(italic lead-in, bold findings, bold-underlined conclusion) of the impression sentence."""
    pres = presentation_text(data).lower()
    lead = (f"Ultrasound obstetric reveals single intrauterine fetus in {pres}{',' if _is_lie_presentation(data) else ''}"
            f" of approx. {_iud_age(data, upper=False)} gestation age with ")
    phrases = ["NO APPRECIABLE CARDIAC ACTIVITY"] + [imp for _b, imp in iud_findings(data)]
    return lead, _iud_join(phrases), " SUGGESTING INTRAUTERINE FETAL DEMISE(IUD)."

def _iud_afi(data):
    """(text before the value, value in bold, abnormal?) for the AFI line; None when nothing was typed."""
    if data.get("afi_method") == "Total 4 Quadrants":
        val = str(data.get("afi_total", "") or "").strip()
        if not val:
            return None
        desc, abn = generate_afi_text(val, is_total=True)
        if desc.startswith("Nil"):
            return desc, "", abn
        return f"{desc[0].lower() + desc[1:]}, total of four quadrants measures ", f"{val}CM", abn
    val = str(data.get("afi_pocket", "") or "").strip()
    if not val:
        return None
    try:
        f = float(val)
        word, abn = (("adequate", False) if 2.0 <= f <= 8.0 else ("reduced", True) if f < 2.0 else ("increased", True))
        return f"{word}, largest pocket measures ", f"{val}CM", abn
    except ValueError:
        return "largest pocket measures ", f"{val}CM", False

def build_iud_docx_bytes(data):
    doc = Document()
    section = doc.sections[0]
    section.page_height, section.page_width = Cm(29.7), Cm(21.0)
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.0)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"; style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0); style.paragraph_format.space_before = Pt(0)

    def _spacer(pt, size=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = Pt(pt)
        r = p.add_run("\u200B"); r.font.size = Pt(size or min(pt, 10))
        return p
    def _line(justify=False):
        p = doc.add_paragraph()
        if justify: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return p

    for _ in range(2): _spacer(12)            # top space for the letterhead (same as the other reports)
    _spacer(6, 5)

    # --- navy title bar + patient box (same as the T16 / T3 reports; bar colour follows the rotation) ---
    _theme = int(data.get("color_theme", ROTATION[0])) % len(COLOR_THEMES)
    add_title_demo_box(doc, data, "USG OBS (INTRAUTERINE FETAL DEMISE)", fill=COLOR_THEMES[_theme][1])
    _spacer(10)
    p = _line(); add_run(p, "LMP- ", bold=True); add_run(p, _iud_lmp(data), italic=True)
    _spacer(8)

    p = _line(justify=True)
    add_run(p, "Uterus shows single fetus. Fetus is in ")
    add_run(p, presentation_text(data).upper(), bold=True)
    add_run(p, " at the time of scanning. Fetal movements are absent. Fetal cardiac pulsations are not appreciated.")
    _spacer(8)

    # --- biometry: plain table, no percentiles ---
    p = _line(); add_run(p, "BIO PHYSICAL PARAMETERS:", bold=True, underline=True, size=10)
    rows = []
    for key, label in (("bpd", "BPD"), ("hc", "HC"), ("ac", "AC"), ("fl", "FL"), ("hl", "HL")):
        size = str(data.get(f"{key}_size", "") or "").strip().lower().replace("mm", "").strip()
        wks = str(data.get(f"{key}_wks", "") or "").strip()
        dys = str(data.get(f"{key}_dys", "") or "").strip()
        if not (size or wks):
            continue
        try: dys_t = f"{int(float(dys)):02d}" if dys else "00"
        except ValueError: dys_t = dys
        rows.append((label, f"{size}mm" if size else "", f"{wks}weeks" if wks else "", f"{dys_t}days\u00b12w" if wks else ""))
    if rows:
        widths = [Cm(2.6), Cm(4.0), Cm(4.0), Cm(4.0)]
        bt = doc.add_table(rows=len(rows), cols=4)
        set_table_borders(bt); set_table_width(bt, widths)
        for i, row in enumerate(rows):
            for j, txt in enumerate(row):
                c = bt.cell(i, j); c.width = widths[j]
                if j: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run(c.paragraphs[0], txt, italic=True, size=10)
    p = _line(); add_run(p, "MEAN AGE :  ", bold=True); add_run(p, _iud_age(data), italic=True)
    _spacer(8)

    # --- AFI / placenta / cervix / spine ---
    afi = _iud_afi(data)
    if afi:
        txt, val, abn = afi
        p = _line(); add_run(p, "AFI:- ", bold=True)
        add_run(p, txt, bold=abn, italic=abn)
        if val: add_run(p, val, bold=True, italic=abn)
    p = _line(); add_run(p, "Placenta ", bold=True); add_run(p, "is ")
    pos, os_ = str(data.get("placenta_pos", "Anterior")).upper(), data.get("placenta_os", "away from the Os")
    add_run(p, pos, bold=True); add_run(p, ", & ")
    if os_ == "Placenta Praevia":
        add_run(p, f"PLACENTA PRAEVIA ({str(data.get('praevia_grade', 'Grade I')).upper()})", bold=True); add_run(p, ".")
    elif os_ == "Low-lying":
        add_run(p, "LOW-LYING", bold=True); add_run(p, lowlying_text(data))
    else:
        add_run(p, "AWAY FROM THE OS", bold=True); add_run(p, ".")
    p = _line(); add_run(p, "CERVIX- ", bold=True)
    cx = fmt_cm(data.get("cervix", "")).replace(" ", "").upper()
    os_open = str(data.get("internal_os", "Closed")).lower() == "open"
    if os_open:
        add_run(p, (f"Length {cx}, " if cx else "")); add_run(p, "INTERNAL OS", bold=True); add_run(p, " is open.")
    else:
        add_run(p, f"Normal ({cx}), " if cx else "Normal, "); add_run(p, "INTERNAL OS", bold=True); add_run(p, " is closed.")
    findings = iud_findings(data)
    if not data.get("iud_spine"):
        p = _line(); add_run(p, "Spine ", bold=True); add_run(p, "appears normal.")
    p = _line(); add_run(p, "\u2192NO CARDIAC ACTIVITY APPRECIATED.", bold=True)
    for body, _imp in findings:
        p = _line(); add_run(p, "\u2192" + body, bold=True)
    _spacer(10)

    # --- impression ---
    p = _line(); add_run(p, "IMPRESSION:", bold=True, underline=True, color=IUD_TEAL)
    lead, found, concl = iud_impression_parts(data)
    p = _line(justify=True)
    add_run(p, lead, italic=True); add_run(p, found, bold=True); add_run(p, concl, bold=True, underline=True)
    _spacer(10)

    # --- PNDT + note box ---
    disc = doc.add_table(rows=1, cols=1)
    disc.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(disc)
    cell = disc.cell(0, 0); p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=8)
    p.add_run().add_break()
    add_run(p, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=8)
    keep_row_together(disc)
    finish_document(doc)
    normalize_table_props(doc)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def iud_preview(data):
    L = ["        ULTRASOUND OBSTETRICS REPORT (IUD)",
         f"NAME :-  {data.get('patient_name', '')}        DATE :-  {data.get('report_date', '')}",
         f"AGE/SEX :-  {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}        REF. BY :-  {data.get('referred_by', '')}",
         "", f"LMP- {_iud_lmp(data)}", "",
         f"Uterus shows single fetus. Fetus is in {presentation_text(data).upper()} at the time of scanning. "
         "Fetal movements are absent. Fetal cardiac pulsations are not appreciated.", "", "BIO PHYSICAL PARAMETERS:"]
    for key, label in (("bpd", "BPD"), ("hc", "HC"), ("ac", "AC"), ("fl", "FL"), ("hl", "HL")):
        if data.get(f"{key}_size") or data.get(f"{key}_wks"):
            _d = str(data.get(f'{key}_dys', '') or '0'); _d = f"{int(float(_d)):02d}" if _d.replace('.', '', 1).isdigit() else _d
            L.append(f"  {label:<4}{str(data.get(f'{key}_size', '')).replace('mm', '')}mm   {data.get(f'{key}_wks', '')}weeks   {_d}days\u00b12w")
    L.append(f"MEAN AGE :  {_iud_age(data)}"); L.append("")
    afi = _iud_afi(data)
    if afi: L.append(f"AFI:- {afi[0]}{afi[1]}")
    L.append(f"Placenta is {str(data.get('placenta_pos', 'Anterior')).upper()}, & {data.get('placenta_os', 'away from the Os').upper()}.")
    L.append(f"CERVIX- {fmt_cm(data.get('cervix', ''))}, INTERNAL OS is {str(data.get('internal_os', 'Closed')).lower()}.")
    if not data.get("iud_spine"): L.append("Spine appears normal.")
    L.append("\u2192NO CARDIAC ACTIVITY APPRECIATED.")
    for body, _imp in iud_findings(data): L.append("\u2192" + body)
    L.append(""); L.append("IMPRESSION:")
    L.append("  " + "".join(iud_impression_parts(data)))
    L.append(""); L.append("PNDT DISCLAIMER: The radiologist has not conducted or caused to conduct any pre-natal diagnostic technique for the purpose of determining fetal sex.")
    return "\n".join(L)


def build_docx_bytes(data, template_id):
    global ACTIVE_THEME
    _an_set_palette("Blue")                # every report starts from the blue palette; the singleton Anomaly Scan may switch it
    if is_twin_template(template_id):
        return build_twins_docx_bytes(data, template_id)
    if template_id == IUD_OBS:            # IUD has its own plain layout
        return build_iud_docx_bytes(data)
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
                           or (template_id == "T3" and bool(doppler_advice(data) or softmarker_advice(data, template_id) or lowlying_advice(data))))   # the advice paragraph needs the room
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
    for _ in range(2):
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

    # 1. TITLE BAR + PATIENT DETAILS BOX (same look as the Anomaly Scan report; each template has its own heading)
    add_title_demo_box(doc, data, REPORT_TITLES.get(template_id, "USG OBS"), fill=COLOR_THEMES[ACTIVE_THEME][1])

    # gap below the box
    p_gap1 = doc.add_paragraph()
    p_gap1.paragraph_format.space_after = Pt(0)
    p_gap1.paragraph_format.line_spacing = Pt((8 if _tight else 10) + (5 if template_id in ("NT_NB", "EARLY_OBS") else 0))

    # 3.5 NT_NB SCAN (laid out like the NT/NB template; shading follows the report's colour theme)
    if template_id in ("NT_NB", "EARLY_OBS"):
        _early = template_id == "EARLY_OBS"        # Early OBS = same look, but its own findings rows, no cranium / DV
        S = early_summary(data) if _early else nt_nb_summary(data)
        _navy = RGBColor(0x1F, 0x4E, 0x79)
        _C = WD_ALIGN_PARAGRAPH.CENTER

        def _nt_head(txt):
            hp = doc.add_paragraph()
            hp.paragraph_format.space_before = Pt(10)
            hp.paragraph_format.space_after = Pt(3)
            hp.paragraph_format.keep_with_next = True
            add_run(hp, txt, bold=True, underline=True, color=_navy)
            return hp

        def _nt_table(rows, widths, head=None):
            """rows: [(label, value)] or [(label, value, pct)]; label column shaded in the theme colour."""
            n = len(rows) + (1 if head else 0)
            t = doc.add_table(rows=n, cols=len(widths))
            set_table_borders(t)
            t.alignment = WD_TABLE_ALIGNMENT.LEFT
            t.autofit = False
            for i, w in enumerate(widths):
                t.columns[i].width = Cm(w)
            for r in t.rows:
                for i, w in enumerate(widths):
                    r.cells[i].width = Cm(w)
            off = 0
            if head:
                for i, h in enumerate(head):
                    set_cell_background(t.cell(0, i), get_gradient_color(0))
                    t.cell(0, i).paragraphs[0].alignment = _C if i else WD_ALIGN_PARAGRAPH.LEFT
                    add_run(t.cell(0, i).paragraphs[0], h, bold=True, size=10)
                off = 1
            for ri, row in enumerate(rows):
                c0 = t.cell(ri + off, 0)
                set_cell_background(c0, get_gradient_color(1 + ri % 2))
                add_run(c0.paragraphs[0], row[0], italic=True, size=10)
                vc = t.cell(ri + off, 1)
                vc.paragraphs[0].alignment = _C
                add_run(vc.paragraphs[0], row[1], italic=True, size=10)
                if len(widths) == 3:
                    pc = t.cell(ri + off, 2)
                    pc.paragraphs[0].alignment = _C
                    _nt_pct_cell(pc.paragraphs[0], row[2] if len(row) > 2 else None)
            return t

        def _nt_pct_cell(p, pct, note=""):
            if pct is None:
                if note:
                    add_run(p, note, italic=True, size=8.5)
                return
            bar = generate_percentile_bar_inline(str(pct))
            add_run(p, f"{pct:g}%  ", bold=True, size=9)
            if bar:
                add_inline_bar(p, bar, width_cm=2.8, text_pt=9)

        lmp = str(data.get('lmp', '') or '').strip()
        if lmp:
            lp = doc.add_paragraph()
            lp.paragraph_format.space_after = Pt(8)
            add_run(lp, "LMP- ", bold=True, italic=True, color=_navy)
            add_run(lp, lmp, italic=True)

        _nt_table(S["early_rows"] if _early else
                  [("CRL", S["crl_line"]), ("Cardiac pulsations", S["cardiac"]),
                   ("Subchorionic collection", S["subchorionic"]), ("Corpus luteal cyst", S["corpus"]),
                   ("Cervix", S["cervix"]), ("Internal Os", S["internal_os"]), ("Placenta", S["placenta"])],
                  [5.5, 11.5])

        if not _early:
            _nt_head("FETAL CRANIUM:")
            ct = _nt_table([("Bilateral choroid Plexi /Falx", S["choroid"], None),
                            ("Nuchal translucency (NT)", S["nt_val"], S["nt_pct"]),
                            ("Nasal bone (NB)", S["nb_val"], S["nb_pct"])],
                           [5.5, 5.5, 6.0], head=["Fetal cranium", "Measurement", "Percentile"])
            _m = ct.cell(1, 1).merge(ct.cell(1, 2))               # choroid row: one wide value cell
            for _extra in _m.paragraphs[1:]:
                _extra._element.getparent().remove(_extra._element)
            if S["nt_pct"] is None and S["nt_pct_note"]:
                _nt_pct_cell(ct.cell(2, 2).paragraphs[0], None, S["nt_pct_note"])

        if S["uta"]:
            _u = S["uta"]
            _nt_head("PRE-ECLAMPSIA SCREEN:")
            up = doc.add_paragraph()
            up.paragraph_format.space_after = Pt(4)
            add_run(up, _u["text"], italic=True)
            _mean = ("Mean Uterine Arteries",) + tuple(_u["mean"][i] or "-" for i in range(3))
            _all = list(_u["rows"]) + [_mean]
            _w = [4.6, 2.3, 2.3, 2.3, 5.5]
            ut = doc.add_table(rows=len(_all) + 1, cols=5)
            set_table_borders(ut)
            ut.alignment = WD_TABLE_ALIGNMENT.LEFT
            ut.autofit = False
            for ci, w in enumerate(_w):
                ut.columns[ci].width = Cm(w)
                for r_ in ut.rows:
                    r_.cells[ci].width = Cm(w)
            for ci, h in enumerate(("Uterine Artery", "PI", "RI", "S/D", "Percentile")):
                set_cell_background(ut.cell(0, ci), get_gradient_color(0))
                ut.cell(0, ci).paragraphs[0].alignment = _C if ci else WD_ALIGN_PARAGRAPH.LEFT
                add_run(ut.cell(0, ci).paragraphs[0], h, bold=True, size=10)
            for ri_, row_ in enumerate(_all, start=1):
                is_mean = ri_ == len(_all)
                set_cell_background(ut.cell(ri_, 0), get_gradient_color(1 + (ri_ - 1) % 2))
                add_run(ut.cell(ri_, 0).paragraphs[0], row_[0], italic=not is_mean, bold=is_mean, size=10)
                for ci in (1, 2, 3):
                    ut.cell(ri_, ci).paragraphs[0].alignment = _C
                    add_run(ut.cell(ri_, ci).paragraphs[0], row_[ci], italic=not is_mean, bold=is_mean, size=10)
                ut.cell(ri_, 4).paragraphs[0].alignment = _C
                if is_mean:
                    _nt_pct_cell(ut.cell(ri_, 4).paragraphs[0], _u["pct"], _u["note"])

        for _head, _d in (("DUCTUS VENOSUS:", None if _early else S["dv"]),):
            if not _d:
                continue
            _nt_head(_head)
            pp = doc.add_paragraph()
            pp.paragraph_format.space_after = Pt(2)
            add_run(pp, _d["text"], italic=True)
            if _d["status"]:
                add_run(pp, f" ({_d['status']})", italic=True)
            add_run(pp, ".  " if _d["pct"] is None else "   ", italic=True)
            if _d["pct"] is not None:
                _bar = generate_percentile_bar_inline(str(_d["pct"]))
                add_run(pp, f"{_d['pct']:g}%  ", bold=True, size=9)
                if _bar:
                    add_inline_bar(pp, _bar, width_cm=2.8, text_pt=11)
            elif _d["note"]:
                add_run(pp, f"({_d['note']})", italic=True, size=9)

        gp = doc.add_paragraph()
        gp.paragraph_format.space_after = Pt(0)
        gp.paragraph_format.line_spacing = Pt(13)
        _edd_rows = [(lab, _fmt_date(data.get(k)), "\u00b12W") for lab, k in (("EDD (LMP)", "edd_lmp"), ("EDD (CUA)", "edd_cua"))
                     if data.get(k) and not (k == "edd_lmp" and lmp == "Not Sure")]
        if _edd_rows:
            et = doc.add_table(rows=len(_edd_rows), cols=3)
            set_table_borders(et)
            et.alignment = WD_TABLE_ALIGNMENT.LEFT
            et.autofit = False
            for ci, w in enumerate((3.0, 3.6, 2.2)):
                et.columns[ci].width = Cm(w)
            for ri, (lab, dt, pm) in enumerate(_edd_rows):
                for ci, w in enumerate((3.0, 3.6, 2.2)):
                    et.cell(ri, ci).width = Cm(w)
                set_cell_background(et.cell(ri, 0), get_gradient_color(1 + ri % 2))
                add_run(et.cell(ri, 0).paragraphs[0], lab, italic=True, size=10)
                et.cell(ri, 1).paragraphs[0].alignment = _C
                add_run(et.cell(ri, 1).paragraphs[0], dt, italic=True, size=10)
                et.cell(ri, 2).paragraphs[0].alignment = _C
                add_run(et.cell(ri, 2).paragraphs[0], pm, italic=True, size=10)

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
        bio_table.alignment = WD_TABLE_ALIGNMENT.CENTER if template_id == COLOR_DOPPLER else WD_TABLE_ALIGNMENT.LEFT
        if template_id == COLOR_DOPPLER:
            set_table_width(bio_table, [Cm(3.4)] * 5)     # full page width, same as the Doppler tables
        
        # 1. Header Row
        headers = ["Fetal growth parameters", "mm", "Weeks", "Days", "Percentile"]
        # Color Doppler: full width. Others: 17 cm = exactly the width of the patient-details box above (right edges line up)
        widths = [Cm(3.4)] * 5 if template_id == COLOR_DOPPLER else [Cm(5.5), Cm(2.0), Cm(2.0), Cm(2.0), Cm(5.5)]
        if template_id != COLOR_DOPPLER:
            set_table_width(bio_table, widths)
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
        if template_id != COLOR_DOPPLER:
            # sits flush under the biometry table, so it has the same 17 cm width (same right edge)
            set_table_width(ratio_table, [Cm(8.5), Cm(8.5)])
            for _c, _w in zip(ratio_table.rows[0].cells, (Cm(8.5), Cm(8.5))):
                _c.width = _w
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
            _add_page_two_break(doc, tight=True)
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
            plac_str = " and Low-lying" + lowlying_text(data)
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

        # --- SOFT MARKERS: last lines under the cord / scar line ---
        for _n in softmarker_notes(data):
            p = doc.add_paragraph(); add_run(p, _n, bold=True, italic=True)

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
    if template_id in ("NT_NB", "EARLY_OBS"):
        _sp = doc.add_paragraph()
        _sp.paragraph_format.space_after = Pt(0)
        _sp.paragraph_format.line_spacing = Pt(5)
    imp_head = doc.add_paragraph()
    add_run(imp_head, "OPINION:" if template_id in ("NT_NB", "EARLY_OBS") else "IMPRESSION:", bold=True, underline=True, color=RGBColor(0x1F, 0x4E, 0x79))
    mean_age = data.get('mean_age', '')
    if template_id == "EARLY_OBS":
        imp_text = early_opinion(data)
    elif template_id == "NT_NB":
        imp_text = nt_nb_opinion(data)
    else:
        pres_text = presentation_text(data)
            
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {pres_text}{',' if _is_lie_presentation(data) else ''} of approx. {mean_age} gestation age"
        _psv_f = doppler_psv_finding(data) if template_id in ("T3", COLOR_DOPPLER) else ""
        _pi_f = doppler_pi_findings(data, capitalise=(template_id == "T3")) if template_id in ("T3", COLOR_DOPPLER) else []
        # T3: pathological PI findings go straight into the impression line; Color Doppler gets its own sentence below.
        # A raised MCA PSV goes into the impression line in both templates.
        _inline = softmarker_phrases(data) + (_pi_f if template_id == "T3" else []) + ([_psv_f] if _psv_f else [])
        imp_text += impression_with_clause(_inline, data)
        imp_text += "."
    if template_id in ("NT_NB", "EARLY_OBS"):
        _sub = doc.add_paragraph()
        _sub.paragraph_format.space_after = Pt(2)
        add_run(_sub, "ULTRASOUND EARLY OBS REVEALS:", size=8)
    p = doc.add_paragraph(); add_run(p, imp_text, bold=True, italic=True)
    _imp_p = p
    _adv_list = doppler_advice_list(data) if template_id in ("T3", COLOR_DOPPLER) else []
    _sm_adv = softmarker_advice(data, template_id) if template_id in ("T16", "T3", COLOR_DOPPLER) else ""
    if _sm_adv:
        _adv_list = _adv_list + [_sm_adv]
    _pl_adv = lowlying_advice(data) if template_id != "EARLY_OBS" else ""
    if _pl_adv:                                 # low-lying placenta reaching / covering the Internal Os
        _adv_list = _adv_list + [_pl_adv]
    _adv = "\n\n".join(_adv_list)
    if template_id == COLOR_DOPPLER and doppler_status(data)["all_normal"]:
        _blank()
        p = doc.add_paragraph(); add_run(p, NORMAL_DOPPLER_LINE, bold=True, italic=True)
    if template_id == COLOR_DOPPLER and _pi_f:    # pathological PI: own sentence after a blank line
        _b = _blank()
        p = doc.add_paragraph(); add_run(p, "Doppler study reveals " + _join_and(_pi_f) + ".", bold=True, italic=True)
        if _adv:
            _b.paragraph_format.keep_with_next = True
            p.paragraph_format.keep_with_next = True
            _imp_p.paragraph_format.keep_with_next = True
    _sf = smaller_fetus_remark(data) if template_id not in ("EARLY_OBS", "NT_NB") else ""
    if _sf:                                    # fetus smaller than LMP dates: remark after a blank line, bold italics
        _b = _blank()
        p = doc.add_paragraph(); add_run(p, _sf, bold=True, italic=True)
        if _adv:                               # keep the chain impression -> remark -> advice -> PNDT box together
            _b.paragraph_format.keep_with_next = True
            p.paragraph_format.keep_with_next = True
            _imp_p.paragraph_format.keep_with_next = True
    if _adv:                                   # pathological Doppler: each advisory after a blank line, bold italics
        for _a in _adv_list:
            _blank().paragraph_format.keep_with_next = True
            p = doc.add_paragraph(); add_run(p, _a, bold=True, italic=True)
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

    keep_row_together(disc_table)
    finish_document(doc)
    if template_id in ("T3", "T16"):
        align_tables_left(doc)         # all left edges on the same line
    normalize_table_props(doc)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ============================================================
# TWIN PREGNANCY TEMPLATE  (Twin A, Twin B, shared uterine arteries, chorionicity / amnionicity)
# ============================================================
def is_twin_template(template_id):
    return str(template_id).startswith("TWINS_")

REPORT_TITLES["TWINS_T3"] = "USG OBS (TWIN PREGNANCY - THIRD TRIMESTER SCAN)"
REPORT_TITLES["TWINS_T16"] = "USG OBS (TWIN PREGNANCY - SECOND TRIMESTER SCAN)"
TWIN_DEFAULT_PRES = {"TWINS_T16": "VARIABLE", "TWINS_T3": "CEPHALIC", "TWINS_DOPPLER": "CEPHALIC", "TWINS_IUD": "CEPHALIC", "TWINS_ANOMALY": "VARIABLE"}
TWINS_IUD = "TWINS_IUD"                                # twin pregnancy with one or both fetuses demised
REPORT_TITLES[TWINS_IUD] = "USG OBS (TWIN PREGNANCY - INTRAUTERINE FETAL DEMISE)"
TWINS_DOPPLER = "TWINS_DOPPLER"                       # Twin Color Doppler template id (always pink, like the singleton one)
REPORT_TITLES[TWINS_DOPPLER] = "USG OBS (TWIN PREGNANCY - COLOR DOPPLER)"
TWIN_NORMAL_DOPPLER_LINE = "Fetal & Maternal vessels reveal normal color flow characteristics in both twins."
TWIN_BIO = [("bpd", "Biparietal Diameter"), ("hc", "Head Circumference"), ("ac", "Abdominal Circumference"),
            ("fl", "Femur Length"), ("hl", "Humerus Length")]
TWIN_PRES_OPTS = ["CEPHALIC", "BREECH", "TRANSVERSE LIE", "OBLIQUE LIE", "VARIABLE"]
TWIN_ANATOMY_DEFAULT = ("Fetal spine is normal at the time of scanning. Fetal skull, gastric bubble, bowel loops, kidneys and "
                        "urinary bladder are normal. No evidence of any abdominal mass is seen. "
                        "No evidence of ascites or pleural effusion is seen.")
TWIN_CERT_OPTS = ["Likely", "Established", "Could not be fairly established"]

def _ga_weeks_float(text):
    g = _split_ga(text)
    return g[0] + g[1] / 7.0 if g else None

def _num_text(v, unit=""):
    s = str(v or "").strip()
    if not s:
        return ""
    try:
        float(s)
        return f"{s} {unit}".strip()
    except ValueError:
        return s

def twin_pres_one(p):
    p = str(p or "CEPHALIC").upper()
    if p in ("TRANSVERSE LIE", "OBLIQUE LIE"):
        return p.capitalize()
    return f"{p.capitalize()} presentation"

def twin_pres_phrase(tw, upper_first=False):
    a, b = tw["A"].get("presentation", "CEPHALIC"), tw["B"].get("presentation", "CEPHALIC")
    if a == b:
        one = twin_pres_one(a)
        s = f"{one} (both twins)" if a in ("TRANSVERSE LIE", "OBLIQUE LIE") else f"{one.replace(' presentation', ' presentations')} (both twins)"
    else:
        s = f"{twin_pres_one(a)} (Twin A) & {twin_pres_one(b)} (Twin B)"
    return s

def twin_chorion_name(data):
    ch = data.get("chorionicity", "Dichorionic")
    am = "Diamniotic" if ch == "Dichorionic" else data.get("amnionicity", "Diamniotic")
    return f"{ch.lower()} {am.lower()}"

def twin_placenta_pos(tw, data, X):
    if data.get("placenta_arr", "Single placenta") == "Single placenta":
        return tw["A"].get("placenta_pos", "Posterior")
    return tw[X].get("placenta_pos", "Posterior")

def twin_chorion_sentence(data, tw):
    am_mono = data.get("chorionicity") == "Monochorionic" and data.get("amnionicity") == "Monoamniotic"
    first = "No intervening amnion seen" if am_mono else "Intervening amnion appreciated"
    if data.get("placenta_arr", "Single placenta") == "Single placenta":
        pl = f"single {twin_placenta_pos(tw, data, 'A').lower()} placenta seen"
    else:
        pl = f"{twin_placenta_pos(tw, data, 'A').lower()} & {twin_placenta_pos(tw, data, 'B').lower()} placentae seen"
    name = twin_chorion_name(data)
    cert = data.get("chorion_cert", "Likely")
    if cert == "Established":
        tail = f"{name.capitalize()} twin pregnancy"
    elif cert == "Likely":
        tail = f"Likely {name} twin pregnancy"
    else:
        tail = f"Chorionicity & amnionicity could not be fairly established at the time of scan \u2013 likely {name} twin pregnancy"
    joiner = "; " if am_mono else " with "
    return f"{first}{joiner}{pl} \u2013 {tail}."

def twin_intro_text(data, tw):
    return (f"Uterus shows twin live fetuses in {twin_pres_phrase(tw)} at the time of scan. "
            f"Fetal movements are present and appear normal. Fetal cardiac pulsations are located.")

def twin_liquor(d):
    """(text, abnormal) for the largest pocket of one twin."""
    val = str(d.get("pocket", "") or "").strip()
    try:
        f = float(val)
    except ValueError:
        return (f"Largest pocket measures {val} cm" if val else "-"), False
    if f == 0:
        return "Anhydramnios, not a single pocket of even one cm seen", True
    if f < 2.0:
        return f"Reduced, largest pocket measures {val} cm", True
    if f > 8.0:
        return f"Increased, largest pocket measures {val} cm", True
    return f"Adequate, largest pocket measures {val} cm", False

def twin_placenta_text(d, pos):
    rel = d.get("placenta_os", "away from the Os")
    det = str(d.get("placenta_detail", "") or "").strip()
    if rel == "Placenta Praevia":
        s = f"{pos}, placenta praevia" + (f" ({det})" if det else "")
    elif rel == "Low-lying":
        clot = "Retroplacental clot is seen." if d.get("clot") else "No retroplacental clot is seen."
        return f"{pos}, low-lying{lowlying_text(d)} {clot}"
    else:
        s = f"{pos}, away from the Os"
    clot = "Retroplacental clot is seen." if d.get("clot") else "No retroplacental clot is seen."
    return f"{s}. {clot}"

def twin_lowlying_advice(tw):
    """One advisory when either twin's placenta reaches / covers the Internal Os (low-lying options 2-4)."""
    return PLAC_LL_ADVICE if any(lowlying_advice(tw[X]) for X in "AB") else ""

def twin_cervix_text(data):
    """'normal (3.1cm), internal Os is closed.' - a typed length gets a cm figure, text such as 'Normal' is used as typed."""
    raw = str(data.get("cervix", "") or "").strip()
    try:
        float(raw)
        head = f"normal ({fmt_cm(raw)})"
    except ValueError:
        head = raw or "normal"
    return f"{head}, internal Os is {str(data.get('internal_os', 'Closed')).lower()}."

def twin_weight_text(d):
    w = str(d.get("fetal_weight", "") or "").strip()
    if not w:
        return "-"
    try:
        fw = float(w)
        return f"{int(fw)} grams \u00b1 {int(round(fw * 0.15))} gms"
    except ValueError:
        return w

def _twin_pct_part(pct):
    return f"; {pct}th percentile" if pct not in (None, "") else ""

def twin_doppler_psv_findings(tw):
    """Raised MCA PSV (>= 1.5 MoM) per twin - goes inside the impression line, as in the singleton Color Doppler report."""
    out = []
    for X in ("A", "B"):
        d = tw[X]
        if doppler_status(d)["psv"] == "Raised":
            out.append(f"raised MCA PSV in Twin {X} ({d.get('mca_psv_mom')} MoM; \u22651.5 MoM), suspicious for fetal anaemia")
    return out

def twin_doppler_pi_findings(data, tw):
    """Pathological PI-type findings for the 'Doppler study reveals ...' sentence (same percentile rules as doppler_status)."""
    out = []
    for X in ("A", "B"):
        d = tw[X]; s = doppler_status(d)
        if s["ua"] == "Raised" and d.get("umbilical_pi"):
            out.append(f"increased Umbilical artery PI in Twin {X} ({d.get('umbilical_pi')}{_twin_pct_part(d.get('umbilical_pct'))})")
        if s["mca"] == "Reduced" and d.get("mca_pi"):
            out.append(f"reduced MCA PI in Twin {X} ({d.get('mca_pi')}{_twin_pct_part(d.get('mca_pct'))})")
        if s["cpr"] == "Reduced" and d.get("cpr"):
            out.append(f"reduced Cerebro-Placental Ratio in Twin {X} ({d.get('cpr')}{_twin_pct_part(d.get('cpr_pct'))})")
    ds = doppler_status(data)
    has_l, has_r = bool(data.get("uterine_left_pi")), bool(data.get("uterine_right_pi"))
    if ds["uta_right"] == "Raised" and data.get("uterine_right_pi"):
        out.append(f"increased Right Uterine artery PI ({data.get('uterine_right_pi')}{_twin_pct_part(data.get('uterine_right_pct'))})")
    if ds["uta_left"] == "Raised" and data.get("uterine_left_pi"):
        out.append(f"increased Left Uterine artery PI ({data.get('uterine_left_pi')}{_twin_pct_part(data.get('uterine_left_pct'))})")
    if ds["uta"] == "Raised" and data.get("uterine_pi") and not (has_l != has_r):      # one side only: the mean just repeats it
        out.append(f"increased Mean Uterine artery PI ({data.get('uterine_pi')}{_twin_pct_part(data.get('uterine_pct'))})")
    if ds["notching"]:
        out.append(f"{str(data.get('uterine_notching')).lower()} in the uterine arteries")
    return out

def twin_doppler_advice_list(data, tw):
    """MCA PSV >= 1.5 MoM in either twin -> fetal-anaemia advisory; any pathological PI finding -> GA-based advisory
    (GA taken from the more advanced twin)."""
    adv = []
    if twin_doppler_psv_findings(tw):
        adv.append(AN_ADVICE_PSV)
    ds = doppler_status(data)
    pathological = (any(doppler_status(tw[X])[k] in ("Raised", "Reduced") for X in ("A", "B") for k in ("ua", "mca", "cpr"))
                    or any(ds[k] == "Raised" for k in ("uta", "uta_left", "uta_right")))
    if pathological:
        adv.append(_advice_for_ga(max((tw["A"], tw["B"]), key=_ga_days)))
    return adv

def twin_doppler_all_normal(data, tw):
    """Both twins have UA and MCA percentiles and the uterine arteries are present - and everything is within range."""
    for X in ("A", "B"):
        s = doppler_status(tw[X])
        if s["ua"] is None or s["mca"] is None:
            return False
        if any(s[k] in ("Raised", "Reduced") for k in ("ua", "mca", "psv", "cpr")):
            return False
    ds = doppler_status(data)
    if ds["uta"] is None or ds["notching"]:
        return False
    return not any(ds[k] == "Raised" for k in ("uta", "uta_left", "uta_right"))

def twin_impression(data, tw, doppler=False):
    name = twin_chorion_name(data)
    pre = "" if data.get("chorion_cert", "Likely") == "Established" else "Likely "
    s = (f"{pre}{name} twin live foetuses of approx. {tw['A'].get('mean_age', '')} & {tw['B'].get('mean_age', '')} "
         f"gestation age, in {twin_pres_phrase(tw)}")
    s = s[0].upper() + s[1:]
    found = []
    for X in ("A", "B"):
        d = tw[X]
        txt, abn = twin_liquor(d)
        if abn:
            kind = "anhydramnios" if txt.startswith("Anhydr") else "oligohydramnios" if txt.startswith("Reduced") else "polyhydramnios"
            found.append(f"{kind} in Twin {X} (largest pocket measures {d.get('pocket')} cm)" if kind != "anhydramnios" else f"anhydramnios in Twin {X}")
        pct = _pf(d.get("umbilical_pct"))
        if not doppler and pct is not None and pct > 95 and d.get("umbilical_pi"):      # Doppler template: own sentence below
            found.append(f"increased umbilical artery PI in Twin {X} ({d.get('umbilical_pi')}; {d.get('umbilical_pct')}th percentile)")
    if doppler:
        found += twin_doppler_psv_findings(tw)
    up = _pf(data.get("uterine_pct"))
    if not doppler and up is not None and up > 95 and data.get("uterine_pi"):
        found.append(f"increased mean uterine artery PI ({data.get('uterine_pi')}; {data.get('uterine_pct')}th percentile)")
    if not doppler and str(data.get("uterine_notching", "No notching")) != "No notching":
        found.append(f"{str(data.get('uterine_notching')).lower()} in the uterine arteries")
    if found:
        s += ", with " + _join_and(found)
    return s + "."

# ---------------------------------------------------------------- twin IUD (one or both twins demised)
def twin_iud_list(tw):
    """Twins marked as demised, e.g. ['B']."""
    return [X for X in "AB" if tw[X].get("iud")]

def twin_iud_findings_text(d):
    """Body text for a demised twin (replaces the anatomy row): spine, cardiac activity and every ticked finding."""
    parts = [] if d.get("iud_spine") else ["SPINE APPEARS NORMAL."]
    parts.append("NO CARDIAC ACTIVITY APPRECIATED.")
    parts += [body for body, _imp in iud_findings(d)]
    return " ".join(parts)

def twin_iud_intro(data, tw):
    dem = twin_iud_list(tw)
    pres = twin_pres_phrase(tw)
    if len(dem) == 2:
        return (f"Uterus shows twin fetuses in {pres} at the time of scan. Fetal movements are absent and "
                f"fetal cardiac pulsations are not appreciated in both twins.")
    if len(dem) == 1:
        d_, l_ = dem[0], "B" if dem[0] == "A" else "A"
        return (f"Uterus shows twin fetuses in {pres} at the time of scan. Twin {l_} is live, with fetal movements present "
                f"and cardiac pulsations located. In Twin {d_}, fetal movements are absent and fetal cardiac pulsations are not appreciated.")
    return twin_intro_text(data, tw)

def twin_iud_view(d):
    """Copy of one twin's data for the report: a demised twin has no percentiles and no heart rate."""
    v_ = dict(d)
    if d.get("iud"):
        for k_ in list(v_):
            if k_.endswith("_pct") and not k_.startswith("iud"):
                v_[k_] = ""
        v_["fetal_weight_pct"] = ""
        v_["fhr"] = "Absent"
    return v_

def twin_iud_extras(data, tw):
    """Other abnormal findings for the impression: liquor in either twin, raised UA PI in a live twin, uterine artery PI / notching."""
    found = []
    for X in "AB":
        d = tw[X]
        txt, abn = twin_liquor(d)
        if abn:
            kind = "anhydramnios" if txt.startswith("Anhydr") else "oligohydramnios" if txt.startswith("Reduced") else "polyhydramnios"
            found.append(f"{kind} in Twin {X}" + ("" if kind == "anhydramnios" else f" (largest pocket measures {d.get('pocket')} cm)"))
        pct = _pf(d.get("umbilical_pct"))
        if not d.get("iud") and pct is not None and pct > 95 and d.get("umbilical_pi"):
            found.append(f"increased umbilical artery PI in Twin {X} ({d.get('umbilical_pi')}; {d.get('umbilical_pct')}th percentile)")
    up = _pf(data.get("uterine_pct"))
    if up is not None and up > 95 and data.get("uterine_pi"):
        found.append(f"increased mean uterine artery PI ({data.get('uterine_pi')}; {data.get('uterine_pct')}th percentile)")
    if str(data.get("uterine_notching", "No notching")) != "No notching":
        found.append(f"{str(data.get('uterine_notching')).lower()} in the uterine arteries")
    return found

def twin_iud_impression_parts(data, tw):
    """(lead-in, bold findings, bold-underlined conclusion, extra sentence) of the twin IUD opinion."""
    dem = twin_iud_list(tw)
    if not dem:
        return twin_impression(data, tw), "", "", ""
    name = twin_chorion_name(data)
    pre = "" if data.get("chorion_cert", "Likely") == "Established" else "Likely "
    lead = (f"{pre}{name} twin pregnancy of approx. {tw['A'].get('mean_age', '')} & {tw['B'].get('mean_age', '')} "
            f"gestation age, in {twin_pres_phrase(tw)}")
    lead = lead[0].upper() + lead[1:]
    def phr(X):
        return _iud_join(["NO APPRECIABLE CARDIAC ACTIVITY"] + [imp for _b, imp in iud_findings(tw[X])])
    if len(dem) == 1:
        d_, l_ = dem[0], "B" if dem[0] == "A" else "A"
        lead += f", with Twin {l_} live and Twin {d_} showing "
        found, concl = phr(d_), f" SUGGESTING INTRAUTERINE FETAL DEMISE(IUD) OF TWIN {d_}."
    else:
        lead += ", with "
        found, concl = f"Twin A showing {phr('A')}; Twin B showing {phr('B')}", " SUGGESTING INTRAUTERINE FETAL DEMISE(IUD) OF BOTH TWINS."
    ex = twin_iud_extras(data, tw)
    return lead, found, concl, (f" Also noted: {_join_and(ex)}." if ex else "")

# ---------------------------------------------------------------- form
def _twin_ga_default(mean_age):
    w = _ga_weeks_float(mean_age)
    return round(min(max(w, 10.0), 42.0), 1) if w else 28.0

def _twin_tab(X, lmp_not_sure, single_placenta, template_id="TWINS_T3"):
    ss = st.session_state
    k = lambda n: f"sr_tw_{X}_{n}"
    d = {}
    if template_id == TWINS_IUD:
        d["iud"] = st.selectbox("Fetal status", ["Live", "IUD (fetal demise)"], key=k("status")) != "Live"
        if d["iud"]:
            st.caption(f"Twin {X}: cardiac activity and movements are reported as absent. Tick what is seen "
                       "(each appears in the body and in the opinion).")
            f1, f2 = st.columns(2)
            for _i, (_fk, _lbl, _b, _imp) in enumerate(IUD_FINDINGS):
                d[f"iud_{_fk}"] = (f1 if _i % 2 == 0 else f2).checkbox(_lbl, key=k(f"iud_{_fk}"))
            d["iud_other"] = st.text_input("Other finding (free text)", key=k("iud_other"))
    b1, b2 = st.columns(2)
    if not ss.get(k("pres_set")):                                   # not chosen by the doctor yet: follow the template default
        ss[k("pres")] = TWIN_DEFAULT_PRES.get(template_id, "CEPHALIC")
    elif k("pres") not in ss and ss.get(k("pres_saved")) in TWIN_PRES_OPTS:
        ss[k("pres")] = ss[k("pres_saved")]                         # Streamlit forgets widgets not shown: restore
    d["presentation"] = b1.selectbox("Presentation", TWIN_PRES_OPTS, key=k("pres"),
                                     on_change=lambda _f=k("pres_set"): ss.__setitem__(_f, True))
    ss[k("pres_saved")] = d["presentation"]
    d["mean_age"] = b2.text_input("Mean age", key=k("mean_age"))
    st.markdown("**Fetal biometry**")
    h = st.columns([2, 1, 1, 1, 3])
    for c, t in zip(h, ("Parameter", "Size (mm)", "Wks", "Days", "Percentile")):
        c.caption(t)
    for p, nm in TWIN_BIO:
        c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 3])
        c1.markdown(f"**{nm}**")
        d[f"{p}_size"] = c2.text_input(f"{X} {p} size", key=k(f"{p}_size"), label_visibility="collapsed")
        d[f"{p}_wks"] = c3.text_input(f"{X} {p} wks", key=k(f"{p}_wks"), label_visibility="collapsed")
        d[f"{p}_dys"] = c4.text_input(f"{X} {p} dys", key=k(f"{p}_dys"), label_visibility="collapsed")
        pv = c5.text_input(f"{X} {p} pct", key=k(f"{p}_pct"), label_visibility="collapsed")
        d[f"{p}_pct"] = "" if lmp_not_sure else pv
        if pv and not lmp_not_sure:
            bar = generate_percentile_bar(pv)
            if bar: c5.image(bar, use_container_width=True)
    b1, b2 = st.columns(2)
    d["fetal_weight"] = b1.text_input("Fetal weight (g)", key=k("weight"))
    wp = b2.text_input("EFW percentile (%)", key=k("weight_pct"))
    d["fetal_weight_pct"] = "" if lmp_not_sure else wp
    if wp and not lmp_not_sure:
        bar = generate_percentile_bar(wp)
        if bar: b2.image(bar, use_container_width=True)
    b1, b2 = st.columns(2)
    d["fhr"] = b1.text_input("FHR (bpm)", key=k("fhr"))
    d["edd_cua"] = b2.text_input("EDD (CUA)", key=k("edd_cua"))
    st.markdown("**Placenta, liquor, anatomy**")
    b1, b2 = st.columns(2)
    if X == "B" and single_placenta:
        b1.caption("Single placenta - same as Twin A")
        d["placenta_pos"] = ""
    else:
        d["placenta_pos"] = b1.selectbox("Placenta position", PLACENTA_OPTS, key=k("plac_pos"))
    d["placenta_os"] = b2.selectbox("Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key=k("plac_os"))
    if d["placenta_os"] == "Placenta Praevia":
        d["placenta_detail"] = st.text_input("Placenta details (optional)", key=k("plac_det"))
    elif d["placenta_os"] == "Low-lying":
        d["placenta_ll"] = lowlying_select(st, k("plac_ll"))
    d["clot"] = st.checkbox("Retroplacental clot present", key=k("clot"))
    d["pocket"] = st.text_input("Largest pocket (cm)", key=k("pocket"))
    _lt, _la = twin_liquor(d)
    if d["pocket"]:
        (st.warning if _la else st.caption)(_lt)
    d["anatomy"] = st.text_area("Anatomy text", value=TWIN_ANATOMY_DEFAULT, key=k("anat"), height=110)
    st.markdown("**Doppler**")
    c1, c2, c3 = st.columns(3)
    d["umbilical_pi"] = c1.text_input("UA PI", key=k("ua_pi"))
    d["umbilical_ri"] = c2.text_input("UA RI", key=k("ua_ri"))
    d["umbilical_sd"] = c3.text_input("UA S/D", key=k("ua_sd"))
    _gd = _twin_ga_default(d["mean_age"])
    d["ga_weeks_dop"] = st.number_input("GA (weeks) for percentiles", min_value=10.0, max_value=42.0, value=_gd, step=0.1,
                                        key=k(f"ga_{_gd}"))
    if d["umbilical_pi"].strip():
        try:
            res = calculate_ua_pi(float(d["umbilical_pi"].replace(",", ".")), d["ga_weeks_dop"])
            if res:
                d["umbilical_pct"] = str(res["percentile"])
                st.caption(f"Umbilical artery: {res['percentile']}th percentile")
                bar = generate_percentile_bar(res["percentile"])
                if bar: st.image(bar, width=220)
        except ValueError:
            st.caption("UA PI must be a number.")
    if template_id == TWINS_DOPPLER:
        st.caption(f"Middle Cerebral Artery - Twin {X}")
        c1, c2, c3, c4 = st.columns(4)
        d["mca_pi"] = c1.text_input("MCA PI", key=k("mca_pi"))
        d["mca_ri"] = c2.text_input("MCA RI", key=k("mca_ri"))
        d["mca_sd"] = c3.text_input("MCA S/D", key=k("mca_sd"))
        d["mca_psv"] = c4.text_input("MCA PSV (cm/s)", key=k("mca_psv"))
        _num = lambda s: float(str(s).replace(",", "."))
        try:
            if d["mca_pi"].strip():
                rm = calculate_mca_pi(_num(d["mca_pi"]), d["ga_weeks_dop"])
                if rm:
                    d["mca_pct"] = str(rm["percentile"])
                    st.caption(f"MCA: {rm['percentile']}th percentile")
                    bar = generate_percentile_bar(rm["percentile"])
                    if bar: st.image(bar, width=220)
                else:
                    st.caption("No MCA percentile below 19 weeks.")
        except ValueError:
            st.caption("MCA PI must be a number.")
        try:                                   # cerebro-placental ratio = MCA PI / UA PI
            if d["mca_pi"].strip() and d["umbilical_pi"].strip():
                _cpr = _num(d["mca_pi"]) / _num(d["umbilical_pi"])
                d["cpr"] = f"{_cpr:.2f}"
                rc = calculate_cpr(_cpr, d["ga_weeks_dop"])
                if rc:
                    d["cpr_pct"] = str(rc["percentile"])
                    st.caption(f"CPR: {d['cpr']} ({rc['percentile']}th percentile)")
                    bar = generate_percentile_bar(rc["percentile"])
                    if bar: st.image(bar, width=220)
        except (ValueError, ZeroDivisionError):
            pass
        try:                                   # MCA peak systolic velocity -> MoM
            if d["mca_psv"].strip():
                rp = calculate_mca_psv_mom(_num(d["mca_psv"]), d["ga_weeks_dop"])
                if rp:
                    d["mca_psv_mom"] = f"{rp['mom']:.2f}"
                    d["mca_psv_median"] = f"{rp['median']:.1f}"
                    _msg = f"MCA PSV: {d['mca_psv']} cm/s = {d['mca_psv_mom']} MoM (median {d['mca_psv_median']} cm/s)"
                    if rp["mom"] >= PSV_MOM_ABNORMAL:
                        st.warning(_msg + " - at or above 1.5 MoM")
                    else:
                        st.caption(_msg)
        except ValueError:
            st.caption("MCA PSV must be a number.")
    d["dv_on"] = st.checkbox("Ductus venosus measured", key=k("dv_on"))
    if d["dv_on"]:
        c1, c2 = st.columns(2)
        d["dv_pi"] = c1.text_input("DV PI", key=k("dv_pi"))
        d["dv_awave"] = c2.selectbox("DV A-wave", ["Positive", "Absent", "Reversed"], key=k("dv_wave"))
        try:
            if d["dv_pi"].strip():
                rdv = calculate_dv_pi(float(d["dv_pi"].replace(",", ".")), d["ga_weeks_dop"])
                if rdv:
                    d["dv_pct"] = str(rdv["percentile"])
                    st.caption(f"Ductus venosus: {rdv['percentile']}th percentile")
                    bar = generate_percentile_bar(rdv["percentile"])
                    if bar: st.image(bar, width=220)
                else:
                    st.caption("No DV percentile outside 20-40 weeks.")
        except ValueError:
            st.caption("DV PI must be a number.")
    return d

def _twin_common_tab(tw, template_id="TWINS_T3"):
    ss = st.session_state
    c = {}
    st.markdown("**Chorionicity & amnionicity**")
    b1, b2 = st.columns(2)
    c["chorionicity"] = b1.radio("Chorionicity", ["Dichorionic", "Monochorionic"], horizontal=True, key="sr_tw_c_chor")
    if c["chorionicity"] == "Monochorionic":
        c["amnionicity"] = b2.radio("Amnionicity", ["Diamniotic", "Monoamniotic"], horizontal=True, key="sr_tw_c_amn")
    else:
        c["amnionicity"] = "Diamniotic"
        b2.caption("Dichorionic twins are always diamniotic.")
    b1, b2 = st.columns(2)
    c["chorion_cert"] = b1.selectbox("How sure", TWIN_CERT_OPTS, key="sr_tw_c_cert")
    c["placenta_arr"] = b2.radio("Placenta", ["Single placenta", "Two placentae"], horizontal=True, key="sr_tw_c_plac")
    if c["chorionicity"] == "Monochorionic" and c["amnionicity"] == "Monoamniotic":
        st.caption("Monoamniotic: the report says no intervening amnion is seen.")
    st.markdown("**Dating, cervix**")
    b1, b2, b3 = st.columns(3)
    c["edd_lmp"] = b1.text_input("EDD (LMP)", key="sr_tw_c_edd_lmp")
    st.session_state.setdefault("sr_tw_c_cervix", "Normal")
    c["cervix"] = b2.text_input("Cervix length (cm) / Normal", key="sr_tw_c_cervix")
    c["internal_os"] = b3.selectbox("Internal Os", ["Closed", "Open"], key="sr_tw_c_os")
    st.markdown("**Uterine arteries (common to both twins)**")
    for side, lab in (("right", "Right"), ("left", "Left")):
        c1, c2, c3 = st.columns(3)
        c[f"uterine_{side}_pi"] = c1.text_input(f"{lab} PI", key=f"sr_tw_c_ut_{side}_pi")
        c[f"uterine_{side}_ri"] = c2.text_input(f"{lab} RI", key=f"sr_tw_c_ut_{side}_ri")
        c[f"uterine_{side}_sd"] = c3.text_input(f"{lab} S/D", key=f"sr_tw_c_ut_{side}_sd")
    def _mean(*vals):
        nums = []
        for x in vals:
            try: nums.append(float(x))
            except (TypeError, ValueError): pass
        return f"{sum(nums) / len(nums):.2f}" if nums else ""
    srd = st.session_state.get("sr_data", {}) or {}
    mpi = _mean(c["uterine_left_pi"], c["uterine_right_pi"]) or str(srd.get("uterine_pi", "") or "")
    mri = _mean(c["uterine_left_ri"], c["uterine_right_ri"]) or str(srd.get("uterine_ri", "") or "")
    msd = _mean(c["uterine_left_sd"], c["uterine_right_sd"]) or str(srd.get("uterine_sd", "") or "")
    c1, c2, c3 = st.columns(3)
    c["uterine_pi"] = c1.text_input("Mean PI", value=mpi, key=f"sr_tw_c_ut_mean_pi_{mpi}")
    c["uterine_ri"] = c2.text_input("Mean RI", value=mri, key=f"sr_tw_c_ut_mean_ri_{mri}")
    c["uterine_sd"] = c3.text_input("Mean S/D", value=msd, key=f"sr_tw_c_ut_mean_sd_{msd}")
    c["uterine_notching"] = st.selectbox("Uterine artery notching", ["No notching", "Unilateral Notching", "Bilateral Notching"], key="sr_tw_c_notch")
    gaa = tw["A"].get("ga_weeks_dop", 28.0)
    if c["uterine_pi"].strip():
        try:
            res = calculate_uta_pi(float(c["uterine_pi"].replace(",", ".")), gaa)
            if res:
                c["uterine_pct"] = str(res["percentile"])
                st.caption(f"Uterine artery (at Twin A's GA, {gaa} w): {res['percentile']}th percentile")
                bar = generate_percentile_bar(res["percentile"])
                if bar: st.image(bar, width=220)
        except ValueError:
            st.caption("Mean PI must be a number.")
    if template_id == TWINS_DOPPLER:           # Right / Left percentiles: same reference curve applied to each side
        for side in ("left", "right"):
            try:
                if c[f"uterine_{side}_pi"].strip():
                    rs = calculate_uta_pi(float(c[f"uterine_{side}_pi"].replace(",", ".")), gaa)
                    if rs:
                        c[f"uterine_{side}_pct"] = str(rs["percentile"])
            except ValueError:
                pass
    return c

TWIN_SR_MAP = (("fetal_weight", "weight"), ("fetal_weight_pct", "weight_pct"), ("fhr", "fhr"), ("mean_age", "mean_age"),
               ("edd_cua", "edd_cua"), ("umbilical_pi", "ua_pi"), ("umbilical_ri", "ua_ri"), ("umbilical_sd", "ua_sd"),
               ("dv_pi", "dv_pi"), ("mca_pi", "mca_pi"), ("mca_ri", "mca_ri"), ("mca_sd", "mca_sd"), ("mca_psv", "mca_psv"),
               ("crl", "crl"), ("nt", "nt"), ("nasal_bone", "nasal_bone"), ("nb_len", "nb_len"), ("nb_pct", "nb_pct"))
TWIN_SR_COMMON = {"edd_lmp": "edd_lmp", "uterine_left_pi": "ut_left_pi", "uterine_left_ri": "ut_left_ri",
                  "uterine_left_sd": "ut_left_sd", "uterine_right_pi": "ut_right_pi", "uterine_right_ri": "ut_right_ri",
                  "uterine_right_sd": "ut_right_sd"}

def _twin_apply_fetus(X, d):
    """Put one fetus's SR values into the Twin X form fields (biometry, weight, FHR, umbilical / MCA / DV Doppler)."""
    ss = st.session_state
    for p, _n in TWIN_BIO:
        if p in d: ss[f"sr_tw_{X}_{p}_size"] = str(d[p])
        for suf in ("wks", "dys", "pct"):
            if f"{p}_{suf}" in d: ss[f"sr_tw_{X}_{p}_{suf}"] = str(d[f"{p}_{suf}"])
    for src, dst in TWIN_SR_MAP:
        if src in d: ss[f"sr_tw_{X}_{dst}"] = str(d[src])
    if "dv_pi" in d: ss[f"sr_tw_{X}_dv_on"] = True
    for p in ("tcd", "ulna", "tibia"):                   # Twin Anomaly: sizes with GA / percentile ...
        if p in d: ss[f"sr_tw_{X}_{p}_size"] = str(d[p])
        for suf in ("wks", "dys", "pct"):
            if f"{p}_{suf}" in d: ss[f"sr_tw_{X}_{p}_{suf}"] = str(d[f"{p}_{suf}"])
    for p in ("lv", "cm", "nf", "radius", "fibula"):     # ... and size-only structures
        if p in d: ss[f"sr_tw_{X}_{p}_size"] = str(d[p])
    for r in ("fl_ac", "fl_bpd", "fl_hc", "hc_ac"):      # machine's own ratios (used when the sizes are not there to work them out)
        if r in d: ss[f"sr_tw_{X}_{r}_sr"] = str(d[r])
    if "crl_wks" in d:                                   # first-trimester GA by CRL
        ss[f"sr_tw_{X}_ga_crl"] = f"{d['crl_wks']}wks {int(d.get('crl_dys', 0)):02d}days"

def twins_form(data, sr_data, v, lmp_not_sure, template_id="TWINS_T3"):
    """Typist form for the twin template. The first DICOM SR of a scan fills Twin A, the next fills Twin B."""
    ss = st.session_state
    ss.setdefault("tw_target", "Twin A")
    if ss.get("tw_applied_v") != v:
        ss["tw_applied_v"] = v
        if sr_data:
            jobs, sequential = [], False
            fetuses = sr_data.get("fetuses")
            n_f = int(sr_data.get("n_fetuses") or 1)
            fid = str(sr_data.get("fetus_id", "") or "")
            if isinstance(fetuses, dict) and len(fetuses) >= 2:        # one SR holding both twins: Fetus 1 -> A, Fetus 2 -> B
                for i, key in enumerate(sorted(fetuses, key=lambda x: (len(x), x))[:2]):
                    jobs.append(("AB"[i], fetuses[key]))
            elif n_f >= 2 and fid:                                        # SR of one numbered fetus of a twin pregnancy
                jobs.append(("A" if fid in ("1", "A") else "B", sr_data))
            else:                                                         # machine sends one unnumbered fetus at a time
                sequential = True
                jobs.append(("A" if ss["tw_target"] == "Twin A" else "B", sr_data))
            for X, d in jobs:
                sig = json.dumps({k: x for k, x in d.items() if k != "fetuses"}, sort_keys=True, default=str)
                if ss.get(f"tw_sig_{X}") != sig:                          # only when this twin's data really changed
                    ss[f"tw_sig_{X}"] = sig                               # (a cumulative SR does not overwrite hand edits)
                    _twin_apply_fetus(X, d)
            csig = json.dumps({k: sr_data.get(k) for k in TWIN_SR_COMMON}, sort_keys=True, default=str)
            if ss.get("tw_sig_c") != csig:
                ss["tw_sig_c"] = csig
                for src, dst in TWIN_SR_COMMON.items():
                    if src in sr_data: ss[f"sr_tw_c_{dst}"] = str(sr_data[src])
            if sequential and jobs[0][0] == "A":
                ss["tw_target"] = "Twin B"      # next unnumbered SR belongs to Twin B
    st.radio("Next unnumbered DICOM SR goes into", ["Twin A", "Twin B"], key="tw_target", horizontal=True)
    st.caption("When the machine numbers the fetuses, Fetus 1 fills Twin A and Fetus 2 fills Twin B automatically. "
               "This switch is only used for SRs that carry no fetus number. Everything stays editable by hand.")
    single = ss.get("sr_tw_c_plac", "Single placenta") == "Single placenta"
    tabs = st.tabs(["Twin A", "Twin B", "Chorionicity & Common"])
    tw = {}
    if template_id in TWIN_FT_TEMPLATES:                 # twin Early OBS / NT-NB
        with tabs[0]: tw["A"] = _twin_ft_tab("A", template_id)
        with tabs[1]: tw["B"] = _twin_ft_tab("B", template_id)
        with tabs[2]: common = _twin_ft_common_tab(template_id)
    elif template_id in TWIN_AN_TEMPLATES:               # twin Anomaly scan
        with tabs[0]: tw["A"] = _twin_an_tab("A", lmp_not_sure, single)
        with tabs[1]: tw["B"] = _twin_an_tab("B", lmp_not_sure, single)
        with tabs[2]: common = _twin_common_tab(tw, template_id)
    else:
        with tabs[0]: tw["A"] = _twin_tab("A", lmp_not_sure, single, template_id)
        with tabs[1]: tw["B"] = _twin_tab("B", lmp_not_sure, single, template_id)
        with tabs[2]: common = _twin_common_tab(tw, template_id)
    data.update(common)
    data["twins"] = tw
    data["mean_age"] = _tw_ga_text(tw["A"])      # dating table / file name use Twin A

# ---------------------------------------------------------------- Word report
def _tw_fix_tables(doc):
    """Every table starts at the left margin with its width set, so the patient box, the tables and the PNDT box line up."""
    st_el = doc.settings.element
    compat = st_el.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat'); st_el.append(compat)
    for cs in compat.findall(qn('w:compatSetting')):
        if cs.get(qn('w:name')) == 'compatibilityMode': compat.remove(cs)
    cs = OxmlElement('w:compatSetting')
    cs.set(qn('w:name'), 'compatibilityMode'); cs.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word'); cs.set(qn('w:val'), '15')
    compat.append(cs)
    for tbl in doc.tables:
        pr = tbl._tbl.tblPr
        grid = [int(g.get(qn('w:w'))) for g in tbl._tbl.tblGrid.findall(qn('w:gridCol'))]
        for tag in ('w:tblW', 'w:jc', 'w:tblInd'):
            for e in pr.findall(qn(tag)): pr.remove(e)
        w = OxmlElement('w:tblW'); w.set(qn('w:type'), 'dxa'); w.set(qn('w:w'), str(sum(grid))); pr.append(w)
        j = OxmlElement('w:jc'); j.set(qn('w:val'), 'left'); pr.append(j)
        i = OxmlElement('w:tblInd'); i.set(qn('w:type'), 'dxa'); i.set(qn('w:w'), '0'); pr.append(i)
        if pr.find(qn('w:tblLayout')) is None:
            lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); pr.append(lay)

def _tw_table(doc, rows, widths):
    t = doc.add_table(rows=rows, cols=len(widths))
    set_table_borders(t)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    for i, w in enumerate(widths):
        t.columns[i].width = Cm(w)
    for r in t.rows:
        for i, w in enumerate(widths):
            r.cells[i].width = Cm(w)
    return t

def _tw_gap(doc, pts):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = Pt(pts)
    return p

def _tw_bar_cell(cell, pct, width_cm=2.6):
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    if pct not in (None, ""):
        bar = generate_percentile_bar_inline(str(pct))
        if bar:
            add_run(p, f"{pct}%  ", bold=True, size=9)
            add_inline_bar(p, bar, width_cm=width_cm, text_pt=9)
            return
    add_run(p, "-", size=10)

def _tw_section_head(doc, txt):
    hp = doc.add_paragraph()
    hp.paragraph_format.space_before = Pt(8); hp.paragraph_format.space_after = Pt(3)
    hp.paragraph_format.keep_with_next = True
    add_run(hp, txt, bold=True, underline=True, color=RGBColor(0x1F, 0x4E, 0x79))

def _tw_banner(doc, X, d):
    t = _tw_table(doc, 1, [9.5, 7.5])
    for c in t.rows[0].cells:
        set_cell_background(c, get_gradient_color(0))
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    lp = t.cell(0, 0).paragraphs[0]
    add_run(lp, f"TWIN {X}", bold=True, size=13, color=RGBColor(0x1F, 0x4E, 0x79))
    if X == "A":
        add_run(lp, "   (PRESENTING TWIN)", bold=True, size=9, color=RGBColor(0x1F, 0x4E, 0x79))
    if d.get("iud"):
        add_run(lp, "   (INTRAUTERINE FETAL DEMISE)", bold=True, size=9, color=RGBColor(0xC0, 0x00, 0x00))
    rp = t.cell(0, 1).paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(rp, twin_pres_one(d.get("presentation")), bold=True, italic=True, size=10)

def _tw_biometry(doc, d):
    t = _tw_table(doc, 6, [5.0, 2.2, 2.2, 2.2, 5.4])
    for i, h in enumerate(("Biometry", "Size (mm)", "Weeks", "Days", "Percentile")):
        c = t.cell(0, i)
        set_cell_background(c, get_gradient_color(0))
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        add_run(c.paragraphs[0], h, bold=True, size=10)
    for ri, (p, nm) in enumerate(TWIN_BIO, start=1):
        bg = get_gradient_color(ri)
        for ci in range(5):
            set_cell_background(t.cell(ri, ci), bg)
        add_run(t.cell(ri, 0).paragraphs[0], nm, size=10)
        for ci, key in ((1, "size"), (2, "wks"), (3, "dys")):
            t.cell(ri, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(t.cell(ri, ci).paragraphs[0], str(d.get(f"{p}_{key}", "") or ""), size=10)
        _tw_bar_cell(t.cell(ri, 4), d.get(f"{p}_pct", ""))

def _tw_summary(doc, d):
    t = _tw_table(doc, 2, [3.2, 5.3, 3.0, 5.5])
    cells = [("Mean age", d.get("mean_age", "") or "-", "FHR", _num_text(d.get("fhr"), "bpm") or "-"),
             ("Fetal weight", twin_weight_text(d), "Percentile", None)]
    for ri, (l1, v1, l2, v2) in enumerate(cells):
        set_cell_background(t.cell(ri, 0), get_gradient_color(1 + ri % 2))
        set_cell_background(t.cell(ri, 2), get_gradient_color(1 + ri % 2))
        add_run(t.cell(ri, 0).paragraphs[0], l1, bold=True, size=10)
        add_run(t.cell(ri, 1).paragraphs[0], v1, italic=True, size=10)
        add_run(t.cell(ri, 2).paragraphs[0], l2, bold=True, size=10)
        t.cell(ri, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.cell(ri, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if v2 is not None:
            add_run(t.cell(ri, 3).paragraphs[0], v2, italic=True, size=10)
        else:
            _tw_bar_cell(t.cell(ri, 3), d.get("fetal_weight_pct", ""))

def _tw_findings(doc, d, pos):
    ltxt, labn = twin_liquor(d)
    rows = [("Placenta", twin_placenta_text(d, pos), False), ("Liquor", ltxt, labn),
            ("Anatomy", str(d.get("anatomy", "") or TWIN_ANATOMY_DEFAULT), False)]
    if d.get("iud"):                       # demised twin: findings in bold instead of the normal anatomy text
        rows[2] = ("Findings", twin_iud_findings_text(d), True)
    t = _tw_table(doc, len(rows), [3.2, 13.8])
    for ri, (lab, txt, abn) in enumerate(rows):
        set_cell_background(t.cell(ri, 0), get_gradient_color(1 + ri % 2))
        add_run(t.cell(ri, 0).paragraphs[0], lab, bold=True, size=10)
        add_run(t.cell(ri, 1).paragraphs[0], txt, bold=abn, italic=True, size=10)

def _tw_doppler_rows(d):
    rows = []
    if any(d.get(k) for k in ("umbilical_pi", "umbilical_ri", "umbilical_sd")):
        rows.append(("Umbilical Artery", d.get("umbilical_pi"), d.get("umbilical_ri"), d.get("umbilical_sd"), d.get("umbilical_pct")))
    if d.get("dv_on") and d.get("dv_pi"):
        rows.append((f"Ductus Venosus (a-wave {str(d.get('dv_awave', 'Positive')).lower()})", d.get("dv_pi"), "", "", d.get("dv_pct")))
    return rows

def _tw_doppler(doc, d):
    rows = _tw_doppler_rows(d)
    if not rows:
        return
    t = _tw_table(doc, 1 + len(rows), [6.2, 1.9, 1.9, 1.9, 5.1])
    for i, h in enumerate(("Doppler", "PI", "RI", "S/D", "Percentile")):
        c = t.cell(0, i)
        set_cell_background(c, get_gradient_color(0))
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        add_run(c.paragraphs[0], h, bold=True, size=10)
    for ri, (lab, pi, rr, sd, pct) in enumerate(rows, start=1):
        set_cell_background(t.cell(ri, 0), get_gradient_color(ri))
        add_run(t.cell(ri, 0).paragraphs[0], lab, size=10)
        for ci, val in ((1, pi), (2, rr), (3, sd)):
            t.cell(ri, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(t.cell(ri, ci).paragraphs[0], str(val or "-"), size=10)
        _tw_bar_cell(t.cell(ri, 4), pct)

# ---- Twin Color Doppler table: Vessel | S/D | RI | PI | Percentile | Interpretation (same layout as the singleton report)
_TW_CD_W = [4.4, 1.3, 1.3, 1.3, 4.6, 4.1]            # 17.0 cm
_TW_OK, _TW_BAD = RGBColor(0x2E, 0x7D, 0x32), RGBColor(0xC6, 0x28, 0x28)

def _tw_interp_cell(cell, text):
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if not text:
        add_run(p, "-", size=10)
    else:
        add_run(p, text, bold=(text != "Normal"), size=10, color=(_TW_OK if text == "Normal" else _TW_BAD))

def _tw_cd_header(t, first="Vessel"):
    for i, h in enumerate((first, "S/D", "RI", "PI", "Percentile", "Interpretation")):
        c = t.cell(0, i)
        set_cell_background(c, get_gradient_color(0))
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        add_run(c.paragraphs[0], h, bold=True, size=10)

def _tw_cd_prep_row(t, ri):
    bg = get_gradient_color(ri)
    for c in t.rows[ri].cells:
        set_cell_background(c, bg)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def _tw_cd_std_row(t, ri, label, sd, rr, pi, pct, interp, sub=None, bold=True):
    _tw_cd_prep_row(t, ri)
    lp = t.cell(ri, 0).paragraphs[0]
    add_run(lp, label, bold=bold, size=10)
    if sub:
        sp = t.cell(ri, 0).add_paragraph(); sp.paragraph_format.space_after = Pt(0)
        add_run(sp, sub, italic=True, size=8)
    for ci, val in ((1, sd), (2, rr), (3, pi)):
        t.cell(ri, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(t.cell(ri, ci).paragraphs[0], str(val) if val else "-", size=10)
    _tw_bar_cell(t.cell(ri, 4), pct, width_cm=2.2)
    _tw_interp_cell(t.cell(ri, 5), interp)

def _tw_cd_span_row(t, ri, label, sub, mid, mid_sub, pct_fn, interp):
    """Row whose value sits across the S/D, RI and PI columns (MCA PSV, cerebro-placental ratio)."""
    _tw_cd_prep_row(t, ri)
    add_run(t.cell(ri, 0).paragraphs[0], label, bold=True, size=10)
    if sub:
        sp = t.cell(ri, 0).add_paragraph(); sp.paragraph_format.space_after = Pt(0)
        add_run(sp, sub, italic=True, size=8)
    span = t.cell(ri, 1).merge(t.cell(ri, 3))
    span.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    span.paragraphs[0].paragraph_format.space_after = Pt(0)
    add_run(span.paragraphs[0], mid, bold=True, size=11)
    if mid_sub:
        sp = span.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER; sp.paragraph_format.space_after = Pt(0)
        add_run(sp, mid_sub, italic=True, size=8)
    pct_fn(t.cell(ri, 4))
    _tw_interp_cell(t.cell(ri, 5), interp)

def _tw_cd_entries(d):
    s = doppler_status(d)
    E = []
    if any(d.get(k) for k in ("umbilical_pi", "umbilical_ri", "umbilical_sd")):
        E.append(("std", "Umbilical Artery", d.get("umbilical_sd"), d.get("umbilical_ri"), d.get("umbilical_pi"),
                  d.get("umbilical_pct"), s["ua"], None))
    if any(d.get(k) for k in ("mca_pi", "mca_ri", "mca_sd")):
        E.append(("std", "Middle Cerebral Artery", d.get("mca_sd"), d.get("mca_ri"), d.get("mca_pi"),
                  d.get("mca_pct"), s["mca"], None))
    if d.get("mca_psv"):
        E.append(("psv", s["psv"]))
    if d.get("cpr"):
        E.append(("cpr", s["cpr"]))
    if d.get("dv_on") and d.get("dv_pi"):
        E.append(("std", "Ductus Venosus", "", "", d.get("dv_pi"), d.get("dv_pct"), None,
                  f"(a-wave {str(d.get('dv_awave', 'Positive')).lower()})"))
    return E

def _tw_color_doppler(doc, d):
    E = _tw_cd_entries(d)
    if not E:
        return
    t = _tw_table(doc, 1 + len(E), _TW_CD_W)
    _tw_cd_header(t)
    for ri, e in enumerate(E, start=1):
        if e[0] == "std":
            _, label, sd, rr, pi, pct, interp, sub = e
            _tw_cd_std_row(t, ri, label, sd, rr, pi, pct, interp, sub=sub)
        elif e[0] == "psv":
            def psv_pct(c, d=d):
                p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if d.get("mca_psv_mom"):
                    p.paragraph_format.space_after = Pt(0)
                    add_run(p, f"{d.get('mca_psv_mom')} MoM", bold=True, size=10)
                    if d.get("mca_psv_median"):
                        p2 = c.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; p2.paragraph_format.space_after = Pt(0)
                        add_run(p2, f"(median {d.get('mca_psv_median')} cm/s)", italic=True, size=8)
                else:
                    add_run(p, "-", size=10)
            _tw_cd_span_row(t, ri, "MCA PSV", "(Peak Systolic Velocity)", f"{d.get('mca_psv')} cm/s", None, psv_pct, e[1])
        else:
            _tw_cd_span_row(t, ri, "Cerebro-Placental Ratio", "(MCA PI \u00f7 UA PI)", f"{d.get('cpr')}", None,
                            lambda c, d=d: _tw_bar_cell(c, d.get("cpr_pct"), width_cm=2.2), e[1])

def _tw_twin_block(doc, X, d, pos, doppler=False):
    G = 6 if doppler else 11        # the Color Doppler twin pages are full: smaller blank lines keep P.T.O on page 1
    """Banner, biometry, then placenta/liquor/anatomy, Doppler (UA, DV - or the full Color Doppler table) and the
    mean age / FHR / weight box. One 11 pt blank line between the biometry and each of the three tables that follow."""
    _tw_banner(doc, X, d)
    _tw_biometry(doc, d)
    _tw_gap(doc, G)
    _tw_findings(doc, d, pos)
    if doppler:
        if _tw_cd_entries(d):
            _tw_gap(doc, G)
            _tw_color_doppler(doc, d)
    elif _tw_doppler_rows(d):
        _tw_gap(doc, G)
        _tw_doppler(doc, d)
    _tw_gap(doc, G)
    _tw_summary(doc, d)

def _tw_dating(doc, data, tw):
    def ga(s):
        g = _split_ga(s)
        return (f"{g[0]} Weeks", f"{g[1]} {'Day' if g[1] == 1 else 'Days'}") if g else ("-", "-")
    rows = []
    lmp = str(data.get("lmp", "") or "").strip()
    if lmp and lmp != "Not Sure":
        wk, dy = ga(data.get("ga_lmp"))
        rows.append(("By LMP", f"LMP: {_fmt_date(lmp)}", wk, dy, _fmt_date(data.get("edd_lmp", "")) or "-"))
    for X in ("A", "B"):
        wk, dy = ga(_tw_ga_text(tw[X]))
        rows.append((f"USG: Twin {X}", "", wk, dy, _fmt_date(tw[X].get("edd_cua", "")) or "-"))
    t = _tw_table(doc, 1 + len(rows), [3.6, 4.6, 2.8, 2.8, 3.2])
    for i, txt in {0: "Dating", 1: "LMP", 4: "EDD"}.items():
        c = t.cell(0, i)
        set_cell_background(c, get_gradient_color(0))
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(c.paragraphs[0], txt, bold=True, size=10)
    gh = t.cell(0, 2).merge(t.cell(0, 3))
    set_cell_background(gh, get_gradient_color(0))
    gh.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(gh.paragraphs[0], "GA", bold=True, size=10)
    for ri, vals in enumerate(rows, start=1):
        for ci, txt in enumerate(vals):
            c = t.cell(ri, ci)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(c.paragraphs[0], txt, size=10)

def _tw_uterine(doc, data):
    if not any(data.get(k) for k in ("uterine_pi", "uterine_left_pi", "uterine_right_pi")):
        return
    _tw_section_head(doc, "UTERINE ARTERY DOPPLER (BOTH TWINS):")
    pct = _pf(data.get("uterine_pct"))
    notch = str(data.get("uterine_notching", "No notching") or "No notching")
    text = (f"{'Raised impedance' if (pct is not None and pct > 95) else 'Normal flow characteristics'} "
            f"seen in the uterine arteries, {notch.lower()} seen.")
    up = doc.add_paragraph(); up.paragraph_format.space_after = Pt(4); add_run(up, text, italic=True)
    rows = []
    for lab, side in (("Rt. Uterine Artery", "uterine_right"), ("Lt. Uterine Artery", "uterine_left")):
        if any(data.get(f"{side}_{m}") for m in ("pi", "ri", "sd")):
            rows.append((lab, data.get(f"{side}_pi"), data.get(f"{side}_ri"), data.get(f"{side}_sd"), False))
    rows.append(("Mean Uterine Arteries", data.get("uterine_pi"), data.get("uterine_ri"), data.get("uterine_sd"), True))
    t = _tw_table(doc, 1 + len(rows), [4.6, 2.3, 2.3, 2.3, 5.5])
    for ci, h in enumerate(("Uterine Artery", "PI", "RI", "S/D", "Percentile")):
        set_cell_background(t.cell(0, ci), get_gradient_color(0))
        t.cell(0, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if ci else WD_ALIGN_PARAGRAPH.LEFT
        add_run(t.cell(0, ci).paragraphs[0], h, bold=True, size=10)
    for ri, (lab, pi, rr, sd, is_mean) in enumerate(rows, start=1):
        set_cell_background(t.cell(ri, 0), get_gradient_color(1 + (ri - 1) % 2))
        add_run(t.cell(ri, 0).paragraphs[0], lab, italic=not is_mean, bold=is_mean, size=10)
        for ci, val in ((1, pi), (2, rr), (3, sd)):
            t.cell(ri, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(t.cell(ri, ci).paragraphs[0], str(val or "-"), italic=not is_mean, bold=is_mean, size=10)
        t.cell(ri, 4).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if is_mean and data.get("uterine_pct"):
            add_run(t.cell(ri, 4).paragraphs[0], f"{data['uterine_pct']}%  ", bold=True, size=9)
            bar = generate_percentile_bar_inline(str(data["uterine_pct"]))
            if bar: add_inline_bar(t.cell(ri, 4).paragraphs[0], bar, width_cm=2.8, text_pt=9)

def _tw_uterine_doppler(doc, data):
    """Uterine arteries (common to both twins) in the Color Doppler layout, with interpretation, then the notching line."""
    keys = [f"{p}_{m}" for p in ("uterine", "uterine_left", "uterine_right") for m in ("pi", "ri", "sd")]
    if not any(data.get(k) for k in keys):
        return
    s = doppler_status(data)
    has_l = any(data.get(f"uterine_left_{m}") for m in ("pi", "ri", "sd"))
    has_r = any(data.get(f"uterine_right_{m}") for m in ("pi", "ri", "sd"))
    if has_l and has_r:
        rows = [("Rt. Uterine Artery", "uterine_right", data.get("uterine_right_pct"), s["uta_right"]),
                ("Lt. Uterine Artery", "uterine_left", data.get("uterine_left_pct"), s["uta_left"]),
                ("Mean Uterine Artery", "uterine", data.get("uterine_pct"), s["uta"])]
    elif has_l or has_r:
        side, label = ("uterine_left", "Lt. Uterine Artery") if has_l else ("uterine_right", "Rt. Uterine Artery")
        rows = [(label, side, data.get("uterine_pct"), s["uta"])]
    else:
        rows = [("Uterine Artery", "uterine", data.get("uterine_pct"), s["uta"])]
    _tw_section_head(doc, "UTERINE ARTERY DOPPLER (BOTH TWINS):")
    t = _tw_table(doc, 1 + len(rows), _TW_CD_W)
    _tw_cd_header(t, "Vessel")
    for ri, (label, pre, pct, interp) in enumerate(rows, start=1):
        _tw_cd_std_row(t, ri, label, data.get(f"{pre}_sd"), data.get(f"{pre}_ri"), data.get(f"{pre}_pi"), pct, interp)
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6)
    add_run(p, "Uterine artery notching: ", bold=True)
    add_run(p, str(data.get("uterine_notching", "No notching")), italic=True)

def _tw_cd_key(doc, tw):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(2)
    add_run(p, DOPPLER_REF_PREFIX, italic=True, size=8)
    add_run(p, "Interpretation key: ", bold=True, size=8)
    add_run(p, "umbilical and uterine artery PI above the 95th percentile, and MCA PI or cerebro-placental ratio "
               "below the 5th percentile, are considered abnormal."
               + (" MCA PSV of 1.5 MoM or more is also considered abnormal (risk of fetal anaemia)."
                  if any(tw[X].get("mca_psv") for X in ("A", "B")) else ""), italic=True, size=8)

def build_twins_docx_bytes(data, template_id="TWINS_T3"):
    if template_id in TWIN_AN_TEMPLATES:
        return build_twins_an_docx_bytes(data)
    if template_id in TWIN_FT_TEMPLATES:
        return build_twins_ft_docx_bytes(data, template_id)
    global ACTIVE_THEME
    ACTIVE_THEME = int(data.get("color_theme", ROTATION[0])) % len(COLOR_THEMES)
    tw = data["twins"]
    is_iud_tw = template_id == TWINS_IUD
    navy = RGBColor(0x1F, 0x4E, 0x79)
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Cm(29.7); sec.page_width = Cm(21.0)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.0)
    nst = doc.styles["Normal"]; nst.font.name = "Calibri"; nst.font.size = Pt(11)
    nst.paragraph_format.space_after = Pt(0); nst.paragraph_format.space_before = Pt(0)
    for pt in (10, 10, 5, 5):
        sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(0); sp.paragraph_format.line_spacing = 1.0
        r = sp.add_run("\u200B"); r.font.name = "Calibri"; r.font.size = Pt(pt)
    add_title_demo_box(doc, data, REPORT_TITLES.get(template_id, "USG OBS (TWIN PREGNANCY)"), fill=COLOR_THEMES[ACTIVE_THEME][1])
    _gp = 6 if template_id == TWINS_DOPPLER else None          # Color Doppler twin pages are full: smaller blank lines
    _tw_gap(doc, _gp or 10)
    lmp = str(data.get("lmp", "") or "").strip()
    if lmp:
        lp = doc.add_paragraph(); lp.paragraph_format.space_after = Pt(4)
        add_run(lp, "LMP- ", bold=True, italic=True, color=navy); add_run(lp, lmp, italic=True)
    ip = doc.add_paragraph(); ip.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; ip.paragraph_format.space_after = Pt(5)
    add_run(ip, twin_iud_intro(data, tw) if is_iud_tw else twin_intro_text(data, tw))
    cb = _tw_table(doc, 1, [17.0])                     # chorionicity / amnionicity call-out
    set_cell_background(cb.cell(0, 0), get_gradient_color(1))
    cp = cb.cell(0, 0).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(cp, twin_chorion_sentence(data, tw), bold=True, italic=True, size=10.5)
    _tw_gap(doc, _gp or 8)
    _tw_dating(doc, data, tw)
    _tw_gap(doc, _gp or 8)
    is_dop = template_id == TWINS_DOPPLER
    _tv = {X: (twin_iud_view(tw[X]) if is_iud_tw else tw[X]) for X in "AB"}      # demised twin: no percentiles / FHR
    _tw_twin_block(doc, "A", _tv["A"], twin_placenta_pos(tw, data, "A"), doppler=is_dop)
    _add_page_two_break(doc, tight=is_dop)
    _tw_twin_block(doc, "B", _tv["B"], twin_placenta_pos(tw, data, "B"), doppler=is_dop)
    if is_dop:
        _tw_uterine_doppler(doc, data)
        _tw_cd_key(doc, tw)
    else:
        _tw_uterine(doc, data)
    cx = doc.add_paragraph(); cx.paragraph_format.space_before = Pt(8)
    add_run(cx, "CERVIX- ", bold=True)
    add_run(cx, twin_cervix_text(data), italic=True)
    _tw_gap(doc, 8)
    ih = doc.add_paragraph(); add_run(ih, "OPINION:", bold=True, underline=True, color=navy)
    sub = doc.add_paragraph(); sub.paragraph_format.space_after = Pt(2); add_run(sub, "OBSTETRIC ULTRASOUND REVEALS:", size=8)
    ipp = doc.add_paragraph()
    if is_iud_tw and twin_iud_list(tw):
        _lead, _found, _concl, _extra = twin_iud_impression_parts(data, tw)
        add_run(ipp, _lead, bold=True, italic=True); add_run(ipp, _found, bold=True)
        add_run(ipp, _concl, bold=True, underline=True)
        if _extra: add_run(ipp, _extra, bold=True, italic=True)
    else:
        add_run(ipp, twin_impression(data, tw, doppler=is_dop), bold=True, italic=True)
    _paras = []                                 # normal-flow line / Doppler findings / advisories, each after a blank line
    if is_dop:
        if twin_doppler_all_normal(data, tw):
            _paras.append(TWIN_NORMAL_DOPPLER_LINE)
        _pi_f = twin_doppler_pi_findings(data, tw)
        if _pi_f:
            _paras.append("Doppler study reveals " + _join_and(_pi_f) + ".")
        _paras += twin_doppler_advice_list(data, tw)
    if twin_lowlying_advice(tw):                # low-lying placenta reaching / covering the Internal Os
        _paras.append(twin_lowlying_advice(tw))
    for _txt in _paras:
        _tw_gap(doc, 11).paragraph_format.keep_with_next = True
        _p = doc.add_paragraph(); add_run(_p, _txt, bold=True, italic=True); _p.paragraph_format.keep_with_next = True
    if _paras:                                  # impression + advice + PNDT box stay together on one page
        ih.paragraph_format.keep_with_next = True; sub.paragraph_format.keep_with_next = True
        ipp.paragraph_format.keep_with_next = True
    _g = _tw_gap(doc, 10)
    if _paras:
        _g.paragraph_format.keep_with_next = True
    dt = doc.add_table(rows=1, cols=1); set_table_borders(dt)
    pc = dt.cell(0, 0).paragraphs[0]; pc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(pc, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=8)
    br = pc.add_run(); br.add_break()
    add_run(pc, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=8)
    keep_row_together(dt)
    finish_document(doc)
    _tw_fix_tables(doc)
    normalize_table_props(doc)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def twins_preview(data, template_id="TWINS_T3"):
    if template_id in TWIN_AN_TEMPLATES:
        return twins_an_preview(data)
    if template_id in TWIN_FT_TEMPLATES:
        return twins_ft_preview(data, template_id)
    tw = data["twins"]
    L = [f"        {REPORT_TITLES.get(template_id, 'USG OBS (TWIN PREGNANCY)')}",
         f"NAME :-  {data.get('patient_name', '')}        DATE :-  {data.get('report_date', '')}",
         f"AGE/SEX :-  {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}        REF. BY :-  {data.get('referred_by', '')}", ""]
    if data.get("lmp"):
        L.append(f"LMP- {data.get('lmp')}")
    L += [twin_iud_intro(data, tw) if template_id == TWINS_IUD else twin_intro_text(data, tw), "", twin_chorion_sentence(data, tw), ""]
    for X in ("A", "B"):
        d = twin_iud_view(tw[X]) if template_id == TWINS_IUD else tw[X]
        L.append(f"TWIN {X}" + ("  (PRESENTING TWIN)" if X == "A" else "") + ("  (INTRAUTERINE FETAL DEMISE)" if d.get("iud") else ""))
        for p, nm in TWIN_BIO:
            L.append(f"  {nm:<26}{str(d.get(p + '_size', '')):>6}  {str(d.get(p + '_wks', '')):>3}w {str(d.get(p + '_dys', '')):>2}d   "
                     f"{(str(d.get(p + '_pct')) + '%') if d.get(p + '_pct') else ''}")
        L.append(f"  Mean age: {d.get('mean_age', '')}    FHR: {_num_text(d.get('fhr'), 'bpm')}")
        L.append(f"  Fetal weight: {twin_weight_text(d)}" + (f"  ({d.get('fetal_weight_pct')}%)" if d.get("fetal_weight_pct") else ""))
        L.append(f"  Placenta: {twin_placenta_text(d, twin_placenta_pos(tw, data, X))}")
        L.append(f"  Liquor: {twin_liquor(d)[0]}")
        L.append(f"  {twin_iud_findings_text(d) if d.get('iud') else d.get('anatomy', '')}")
        if d.get("umbilical_pi"):
            L.append(f"  Umbilical artery PI {d.get('umbilical_pi')}" + (f" ({d.get('umbilical_pct')}th percentile)" if d.get("umbilical_pct") else ""))
        if d.get("mca_pi"):
            L.append(f"  MCA PI {d.get('mca_pi')}" + (f" ({d.get('mca_pct')}th percentile)" if d.get("mca_pct") else ""))
        if d.get("mca_psv"):
            L.append(f"  MCA PSV {d.get('mca_psv')} cm/s" + (f" = {d.get('mca_psv_mom')} MoM" if d.get("mca_psv_mom") else ""))
        if d.get("cpr"):
            L.append(f"  Cerebro-placental ratio {d.get('cpr')}" + (f" ({d.get('cpr_pct')}th percentile)" if d.get("cpr_pct") else ""))
        if d.get("dv_on") and d.get("dv_pi"):
            L.append(f"  Ductus venosus PI {d.get('dv_pi')} (a-wave {str(d.get('dv_awave', 'Positive')).lower()})"
                     + (f" ({d.get('dv_pct')}th percentile)" if d.get("dv_pct") else ""))
        L.append("")
    if data.get("uterine_pi"):
        L.append(f"Mean uterine artery PI {data.get('uterine_pi')}" + (f" ({data.get('uterine_pct')}th percentile)" if data.get("uterine_pct") else "")
                 + f", {str(data.get('uterine_notching', 'No notching')).lower()}")
    L.append("CERVIX- " + twin_cervix_text(data))
    is_dop = template_id == TWINS_DOPPLER
    L += ["", "OPINION:", ("".join(twin_iud_impression_parts(data, tw)) if template_id == TWINS_IUD and twin_iud_list(tw)
                      else twin_impression(data, tw, doppler=is_dop))]
    if is_dop:
        if twin_doppler_all_normal(data, tw):
            L += ["", TWIN_NORMAL_DOPPLER_LINE]
        _pi_f = twin_doppler_pi_findings(data, tw)
        if _pi_f:
            L += ["", "Doppler study reveals " + _join_and(_pi_f) + "."]
        for _a in twin_doppler_advice_list(data, tw):
            L += ["", _a]
    if twin_lowlying_advice(tw):
        L += ["", twin_lowlying_advice(tw)]
    return "\n".join(L)


# ============================================================
# TWIN EARLY OBS and TWIN NT / NB TEMPLATES
# ============================================================
TWINS_EARLY = "TWINS_EARLY"
TWINS_NT_NB = "TWINS_NT_NB"
TWIN_FT_TEMPLATES = (TWINS_EARLY, TWINS_NT_NB)
REPORT_TITLES[TWINS_EARLY] = "USG OBS (TWIN EARLY PREGNANCY SCAN)"
REPORT_TITLES[TWINS_NT_NB] = "USG OBS (TWIN NT / NB SCAN)"
TWIN_CARDIAC_OPTS = ["Present", "Not yet appreciable", "Absent"]
TWIN_FT_PLACENTA_OPTS = ["-"] + PLACENTA_OPTS            # "-" = placenta not yet formed: the report shows " - "

def _tw_ga_text(d):
    return str(d.get("ga_crl") or d.get("mean_age") or "").strip()

def _tw_flat(data, d, with_uterine=True):
    """One flat dict (common data + one twin) for the singleton nt_nb_summary(); '-' placenta = not formed."""
    m = {k: v for k, v in data.items() if k != "twins"}
    m.update(d)
    if str(m.get("placenta_pos", "")).strip() in ("-", ""):
        m["placenta_pos"] = ""
    if not with_uterine:
        for k in [k for k in m if k.startswith("uterine")]:
            m.pop(k)
    return m

def twin_ft_summaries(data, tw):
    """Per-twin summaries without the uterine arteries, plus a common one (Twin A's CRL) holding the uterine arteries."""
    SA = nt_nb_summary(_tw_flat(data, tw["A"], False))
    SB = nt_nb_summary(_tw_flat(data, tw["B"], False))
    SC = nt_nb_summary(_tw_flat(data, tw["A"], True))
    return {"A": SA, "B": SB, "C": SC}

def twin_ft_cardiac(d):
    st_ = d.get("cardiac_status", "Present")
    fhr = str(d.get("fhr_val", "") or "").strip()
    if st_ == "Present":
        return "Present" + (f", {fhr} bpm" if fhr else "")
    return "Not yet appreciable" if st_ == "Not yet appreciable" else "Not appreciated"

def twin_ft_chorion_pair(data):
    ch = data.get("chorionicity", "Dichorionic")
    am = "Diamniotic" if ch == "Dichorionic" else data.get("amnionicity", "Diamniotic")
    return ch, am

def twin_ft_sentence(data):
    ch, am = twin_ft_chorion_pair(data)
    name = f"{ch.lower()} and {am.lower()}"
    sign = ("amnion & lambda sign evident" if ch == "Dichorionic"
            else "no intervening amnion seen" if am == "Monoamniotic" else "amnion & T sign evident")
    cert = data.get("chorion_cert", "Established")
    if cert == "Established":
        return f"{name.capitalize()} pregnancy ({sign})."
    if cert == "Likely":
        return f"Likely {name} pregnancy ({sign})."
    return f"Chorionicity & amnionicity could not be fairly established at the time of scan \u2013 likely {name} pregnancy."

def twin_ft_impression(data, tw, nt_nb):
    ch, am = twin_ft_chorion_pair(data)
    cert = data.get("chorion_cert", "Established")
    name = f"{ch} and {am}" if cert == "Established" else f"Likely {ch.lower()} and {am.lower()}"
    ga = {X: _tw_ga_text(tw[X]) for X in "AB"}
    s = f"{name} intra uterine twin live pregnancies of {ga['A']} +/-1w & {ga['B']} +/-1w"
    present = [X for X in "AB" if tw[X].get("cardiac_status", "Present") == "Present"]
    absent = [X for X in "AB" if X not in present]
    def _why(X):
        return ("cardiac activity not yet appreciable" if tw[X].get("cardiac_status") == "Not yet appreciable"
                else "no cardiac activity appreciated")
    if not absent:
        s += " with good cardiac activity" + (" in both the fetuses" if nt_nb else "")
    elif not present:
        s += (" with cardiac activity not yet appreciable in both the fetuses"
              if all(tw[X].get("cardiac_status") == "Not yet appreciable" for X in "AB")
              else " with no cardiac activity appreciated in both the fetuses")
    else:
        s += f" with good cardiac activity in Twin {present[0]} and {_why(absent[0])} in Twin {absent[0]}"
    S = twin_ft_summaries(data, tw)
    flags = []
    if nt_nb:
        for X in "AB":
            flags += [f"{f} in Twin {X}" for f in S[X]["flags"]]
    flags += [f for f in S["C"]["flags"] if "uterine" in f]
    if flags:
        s += ", with " + _join_and(flags)
    return s + "."

# ---------------------------------------------------------------- form
def _twin_ft_tab(X, template_id):
    k = lambda n: f"sr_tw_{X}_{n}"
    nt_nb = template_id == TWINS_NT_NB
    d = {}
    c1, c2 = st.columns(2)
    d["crl"] = c1.text_input("CRL (mm)", key=k("crl"))
    d["ga_crl"] = c2.text_input("GA by CRL", key=k("ga_crl"))
    c1, c2 = st.columns(2)
    d["cardiac_status"] = c1.selectbox("Cardiac pulsations", TWIN_CARDIAC_OPTS, key=k("cardiac"))
    d["fhr_val"] = c2.text_input("FHR (bpm)", key=k("fhr")) if d["cardiac_status"] == "Present" else ""
    if not nt_nb:
        d["yolk"] = st.text_input("Yolk sac", value="Not Seen", key=k("yolk"))
    c1, c2 = st.columns(2)
    d["subchorionic_present"] = c1.checkbox("Subchorionic collection present", key=k("sch"))
    if d["subchorionic_present"]:
        d["subchorionic_grade"] = c2.radio("Grade", ["Grade-I", "Grade-I/II", "Grade-II", "Grade-II/III"], horizontal=True, key=k("sch_g"))
    c1, c2 = st.columns(2)
    d["placenta_pos"] = c1.selectbox("Placenta", TWIN_FT_PLACENTA_OPTS,
                                     format_func=lambda o: "Not formed yet (shows  -  )" if o == "-" else o, key=k("plac_ft"))
    if d["placenta_pos"] != "-":
        d["placenta_os"] = c2.selectbox("Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key=k("plac_os_ft"))
        if d["placenta_os"] == "Low-lying":
            d["placenta_ll"] = lowlying_select(st, k("plac_ll_ft"))
    d["edd_cua"] = st.text_input("EDD (CUA)", key=k("edd_cua"))
    if nt_nb:
        st.markdown("**Fetal cranium**")
        c1, c2 = st.columns(2)
        d["choroid"] = c1.selectbox("Choroid plexi / Falx", ["Appreciated, appear Normal", "Not well visualised"], key=k("choroid"))
        d["nt"] = c2.text_input("NT (mm)", key=k("nt"))
        c1, c2, c3 = st.columns(3)
        d["nasal_bone"] = c1.text_input("Nasal bone (Present / Absent)", key=k("nasal_bone"))
        d["nb_len"] = c2.text_input("NB length (mm)", key=k("nb_len"))
        d["nb_pct"] = c3.text_input("NB percentile (machine)", key=k("nb_pct"))
        st.markdown("**Ductus venosus (optional)**")
        c1, c2 = st.columns(2)
        d["dv_pi"] = c1.text_input("DV PI", key=k("dv_pi"))
        d["dv_awave"] = c2.selectbox("DV A-wave", ["Positive", "Absent", "Reversed"], key=k("dv_wave"))
    return d

def _twin_ft_common_tab(template_id):
    c = {}
    st.markdown("**Chorionicity & amnionicity**")
    b1, b2 = st.columns(2)
    c["chorionicity"] = b1.radio("Chorionicity", ["Dichorionic", "Monochorionic"], horizontal=True, key="sr_tw_c_chor")
    if c["chorionicity"] == "Monochorionic":
        c["amnionicity"] = b2.radio("Amnionicity", ["Diamniotic", "Monoamniotic"], horizontal=True, key="sr_tw_c_amn")
    else:
        c["amnionicity"] = "Diamniotic"
        b2.caption("Dichorionic twins are always diamniotic.")
    c["chorion_cert"] = st.selectbox("How sure", ["Established", "Likely", "Could not be fairly established"], key="sr_tw_c_cert_ft")
    st.markdown("**Common features**")
    b1, b2, b3 = st.columns(3)
    st.session_state.setdefault("sr_tw_c_cervix", "Normal")
    c["cervix"] = b1.text_input("Cervix length (cm) / Normal", key="sr_tw_c_cervix")
    c["internal_os"] = b2.selectbox("Internal Os", ["Closed", "Open"], key="sr_tw_c_os")
    c["edd_lmp"] = b3.text_input("EDD (LMP)", key="sr_tw_c_edd_lmp")
    b1, b2 = st.columns(2)
    c["corpus_present"] = b1.checkbox("Corpus luteal cyst present", key="sr_tw_c_cl")
    if c["corpus_present"]:
        d1, d2 = b2.columns(2)
        c["corpus_side"] = d1.selectbox("Side", ["Right", "Left"], key="sr_tw_c_cl_side")
        c["corpus_size"] = d2.text_input("Size (mm)", key="sr_tw_c_cl_size")
    st.markdown("**Uterine arteries - pre-eclampsia screen (box appears only when values are entered)**")
    for side, lab in (("right", "Right"), ("left", "Left")):
        c1, c2, c3 = st.columns(3)
        c[f"uterine_{side}_pi"] = c1.text_input(f"{lab} PI", key=f"sr_tw_c_ut_{side}_pi")
        c[f"uterine_{side}_ri"] = c2.text_input(f"{lab} RI", key=f"sr_tw_c_ut_{side}_ri")
        c[f"uterine_{side}_sd"] = c3.text_input(f"{lab} S/D", key=f"sr_tw_c_ut_{side}_sd")
    def _mean(*vals):
        nums = []
        for x in vals:
            try: nums.append(float(x))
            except (TypeError, ValueError): pass
        return f"{sum(nums) / len(nums):.2f}" if nums else ""
    srd = st.session_state.get("sr_data", {}) or {}
    mpi = _mean(c["uterine_left_pi"], c["uterine_right_pi"]) or str(srd.get("uterine_pi", "") or "")
    mri = _mean(c["uterine_left_ri"], c["uterine_right_ri"]) or str(srd.get("uterine_ri", "") or "")
    msd = _mean(c["uterine_left_sd"], c["uterine_right_sd"]) or str(srd.get("uterine_sd", "") or "")
    c1, c2, c3 = st.columns(3)
    c["uterine_pi"] = c1.text_input("Mean PI", value=mpi, key=f"sr_tw_c_ut_mean_pi_{mpi}")
    c["uterine_ri"] = c2.text_input("Mean RI", value=mri, key=f"sr_tw_c_ut_mean_ri_{mri}")
    c["uterine_sd"] = c3.text_input("Mean S/D", value=msd, key=f"sr_tw_c_ut_mean_sd_{msd}")
    c1, c2 = st.columns(2)
    c["uterine_notching"] = c1.selectbox("Uterine artery notching", ["No notching", "Unilateral notching", "Bilateral notching"], key="sr_tw_c_notch")
    c["conception"] = c2.selectbox("Conception (UtA PI reference)", ["Natural", "IVF"], key="sr_tw_c_conception")
    st.caption("Uterine artery percentile is worked out from Twin A's CRL (valid for CRL 45-84 mm).")
    return c

# ---------------------------------------------------------------- Word report
def _tw_kv_table(doc, rows, widths=(5.5, 11.5)):
    t = _tw_table(doc, len(rows), list(widths))
    for ri, (lab, val) in enumerate(rows):
        set_cell_background(t.cell(ri, 0), get_gradient_color(1 + ri % 2))
        add_run(t.cell(ri, 0).paragraphs[0], lab, italic=True, size=10)
        t.cell(ri, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(t.cell(ri, 1).paragraphs[0], val, italic=True, size=10)
    return t

def _tw_ft_pct_cell(p, pct, note=""):
    if pct is None:
        if note:
            add_run(p, note, italic=True, size=8.5)
        return
    bar = generate_percentile_bar_inline(str(pct))
    add_run(p, f"{pct:g}%  ", bold=True, size=9)
    if bar:
        add_inline_bar(p, bar, width_cm=2.8, text_pt=9)

def _tw_ft_banner(doc, X, right):
    t = _tw_table(doc, 1, [9.5, 7.5])
    for c in t.rows[0].cells:
        set_cell_background(c, get_gradient_color(0))
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    add_run(t.cell(0, 0).paragraphs[0], f"TWIN {X}", bold=True, size=13, color=RGBColor(0x1F, 0x4E, 0x79))
    rp = t.cell(0, 1).paragraphs[0]; rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(rp, right, bold=True, italic=True, size=10)

def _tw_ft_findings(doc, X, d, S, nt_nb):
    rows = [("CRL", S["crl_line"] or "-")]
    if not nt_nb:
        rows.append(("Yolk sac", str(d.get("yolk", "") or "Not Seen")))
    rows += [("Cardiac pulsations", twin_ft_cardiac(d)), ("Subchorionic collection", S["subchorionic"]),
             ("Placenta", S["placenta"])]
    _tw_ft_banner(doc, X, _tw_ga_text(d))
    _tw_kv_table(doc, rows)

def _tw_ft_cranium(doc, X, S):
    _tw_section_head(doc, f"FETAL CRANIUM ({X}):")
    rows = [("Bilateral choroid Plexi /Falx", S["choroid"], None), ("Nuchal translucency (NT)", S["nt_val"], S["nt_pct"]),
            ("Nasal bone (NB)", S["nb_val"], S["nb_pct"])]
    t = _tw_table(doc, 4, [5.5, 5.5, 6.0])
    for i, h in enumerate(("Fetal cranium", "Measurement", "Percentile")):
        set_cell_background(t.cell(0, i), get_gradient_color(0))
        t.cell(0, i).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
        add_run(t.cell(0, i).paragraphs[0], h, bold=True, size=10)
    for ri, (lab, val, pct) in enumerate(rows, start=1):
        set_cell_background(t.cell(ri, 0), get_gradient_color(1 + (ri - 1) % 2))
        add_run(t.cell(ri, 0).paragraphs[0], lab, italic=True, size=10)
        t.cell(ri, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(t.cell(ri, 1).paragraphs[0], val, italic=True, size=10)
        t.cell(ri, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if ri > 1:
            _tw_ft_pct_cell(t.cell(ri, 2).paragraphs[0], pct, S["nt_pct_note"] if (ri == 2 and pct is None) else "")
    m = t.cell(1, 1).merge(t.cell(1, 2))                    # choroid row: one wide value cell
    for extra in m.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)

def _tw_ft_dv(doc, X, S):
    dv = S["dv"]
    if not dv:
        return
    _tw_section_head(doc, f"DUCTUS VENOSUS ({X}):")
    pp = doc.add_paragraph(); pp.paragraph_format.space_after = Pt(2)
    add_run(pp, dv["text"], italic=True)
    if dv["status"]:
        add_run(pp, f" ({dv['status']})", italic=True)
    add_run(pp, ".  " if dv["pct"] is None else "   ", italic=True)
    if dv["pct"] is not None:
        bar = generate_percentile_bar_inline(str(dv["pct"]))
        add_run(pp, f"{dv['pct']:g}%  ", bold=True, size=9)
        if bar:
            add_inline_bar(pp, bar, width_cm=2.8, text_pt=11)
    elif dv["note"]:
        add_run(pp, f"({dv['note']})", italic=True, size=9)

def _tw_ft_common_table(doc, S):
    t = _tw_table(doc, 1, [17.0])
    set_cell_background(t.cell(0, 0), get_gradient_color(0))
    add_run(t.cell(0, 0).paragraphs[0], "COMMON FEATURES", bold=True, size=11, color=RGBColor(0x1F, 0x4E, 0x79))
    _tw_kv_table(doc, [("Corpus luteal cyst", S["corpus"]), ("Cervix", S["cervix"]), ("Internal Os", S["internal_os"])])

def _tw_ft_uterine(doc, S):
    u = S["uta"]
    if not u:                                   # no uterine artery values -> no pre-eclampsia screen box
        return
    _tw_section_head(doc, "PRE-ECLAMPSIA SCREEN:")
    up = doc.add_paragraph(); up.paragraph_format.space_after = Pt(4); add_run(up, u["text"], italic=True)
    mean = ("Mean Uterine Arteries",) + tuple(u["mean"][i] or "-" for i in range(3))
    allr = list(u["rows"]) + [mean]
    t = _tw_table(doc, len(allr) + 1, [4.6, 2.3, 2.3, 2.3, 5.5])
    for ci, h in enumerate(("Uterine Artery", "PI", "RI", "S/D", "Percentile")):
        set_cell_background(t.cell(0, ci), get_gradient_color(0))
        t.cell(0, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if ci else WD_ALIGN_PARAGRAPH.LEFT
        add_run(t.cell(0, ci).paragraphs[0], h, bold=True, size=10)
    for ri, row in enumerate(allr, start=1):
        is_mean = ri == len(allr)
        set_cell_background(t.cell(ri, 0), get_gradient_color(1 + (ri - 1) % 2))
        add_run(t.cell(ri, 0).paragraphs[0], row[0], italic=not is_mean, bold=is_mean, size=10)
        for ci in (1, 2, 3):
            t.cell(ri, ci).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(t.cell(ri, ci).paragraphs[0], row[ci], italic=not is_mean, bold=is_mean, size=10)
        t.cell(ri, 4).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if is_mean:
            _tw_ft_pct_cell(t.cell(ri, 4).paragraphs[0], u["pct"], u["note"])

def build_twins_ft_docx_bytes(data, template_id):
    global ACTIVE_THEME
    nt_nb = template_id == TWINS_NT_NB
    ACTIVE_THEME = PINK_THEME_INDEX if nt_nb else int(data.get("color_theme", ROTATION[0])) % len(COLOR_THEMES)
    tw = data["twins"]
    S = twin_ft_summaries(data, tw)
    navy = RGBColor(0x1F, 0x4E, 0x79)
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Cm(29.7); sec.page_width = Cm(21.0)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2.0)
    nst = doc.styles["Normal"]; nst.font.name = "Calibri"; nst.font.size = Pt(11)
    nst.paragraph_format.space_after = Pt(0); nst.paragraph_format.space_before = Pt(0)
    for pt in (10, 10, 5, 5):
        sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(0); sp.paragraph_format.line_spacing = 1.0
        r = sp.add_run("\u200B"); r.font.name = "Calibri"; r.font.size = Pt(pt)
    add_title_demo_box(doc, data, REPORT_TITLES[template_id], fill=COLOR_THEMES[ACTIVE_THEME][1])
    _tw_gap(doc, 6)
    lmp = str(data.get("lmp", "") or "").strip()
    if lmp:
        lp = doc.add_paragraph(); lp.paragraph_format.space_after = Pt(4)
        add_run(lp, "LMP- ", bold=True, italic=True, color=navy); add_run(lp, lmp, italic=True)
    cb = _tw_table(doc, 1, [17.0])                     # chorionicity / amnionicity call-out
    set_cell_background(cb.cell(0, 0), get_gradient_color(1))
    cp = cb.cell(0, 0).paragraphs[0]; cp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(cp, twin_ft_sentence(data), bold=True, italic=True, size=10.5)
    _tw_gap(doc, 5)
    for X in "AB":
        _tw_ft_findings(doc, X, tw[X], S[X], nt_nb)
        if nt_nb:
            _tw_gap(doc, 4)
            _tw_ft_cranium(doc, X, S[X])
            _tw_ft_dv(doc, X, S[X])
        _tw_gap(doc, 9 if nt_nb else 5)
    if nt_nb:
        _add_page_two_break(doc)
    _tw_ft_common_table(doc, S["C"])
    _tw_gap(doc, 2)
    _tw_ft_uterine(doc, S["C"])
    _tw_gap(doc, 5)
    _tw_dating(doc, data, tw)
    _tw_gap(doc, 5)
    ih = doc.add_paragraph(); add_run(ih, "OPINION:", bold=True, underline=True, color=navy)
    sub = doc.add_paragraph(); sub.paragraph_format.space_after = Pt(2); add_run(sub, "ULTRASOUND EARLY OBS REVEALS:", size=8)
    ipp = doc.add_paragraph(); add_run(ipp, twin_ft_impression(data, tw, nt_nb), bold=True, italic=True)
    if twin_lowlying_advice(tw):                # low-lying placenta reaching / covering the Internal Os
        ih.paragraph_format.keep_with_next = True; sub.paragraph_format.keep_with_next = True
        ipp.paragraph_format.keep_with_next = True
        _tw_gap(doc, 11).paragraph_format.keep_with_next = True
        _ap = doc.add_paragraph(); add_run(_ap, twin_lowlying_advice(tw), bold=True, italic=True)
        _ap.paragraph_format.keep_with_next = True
    _g = _tw_gap(doc, 6)
    if twin_lowlying_advice(tw):
        _g.paragraph_format.keep_with_next = True      # advisory + PNDT box stay together
    dt = doc.add_table(rows=1, cols=1); set_table_borders(dt)
    pc = dt.cell(0, 0).paragraphs[0]; pc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(pc, '"The radiologist has not conducted or caused to conduct any pre natal diagnostic technique (sonography) for the purpose of determining the sex of the foetus. No person including the person conducting the diagnostic procedure has communicated to the pregnant woman concerned or her relatives or any other person the sex of the foetus by words, signs, or in any other manner."', bold=True, italic=True, size=8)
    br = pc.add_run(); br.add_break()
    add_run(pc, "Note:- This is a obstetrical ultrasound, mainly done for estimation of gestational age, amount of liquor, placental position and general well-being of the fetus. Moreover, the anomalies in relation to fetal heart and limbs are extremely difficult to assess due to constantly changing position of the fetus and overlapping of its various parts.", bold=True, italic=True, size=8)
    keep_row_together(dt)
    finish_document(doc)
    _tw_fix_tables(doc)
    normalize_table_props(doc)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def twins_ft_preview(data, template_id):
    nt_nb = template_id == TWINS_NT_NB
    tw = data["twins"]; S = twin_ft_summaries(data, tw)
    L = [f"        {REPORT_TITLES[template_id]}",
         f"NAME :-  {data.get('patient_name', '')}        DATE :-  {data.get('report_date', '')}",
         f"AGE/SEX :-  {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}        REF. BY :-  {data.get('referred_by', '')}", ""]
    if data.get("lmp"):
        L.append(f"LMP- {data.get('lmp')}")
    L += [twin_ft_sentence(data), ""]
    for X in "AB":
        d, s = tw[X], S[X]
        L.append(f"TWIN {X}   {_tw_ga_text(d)}")
        L.append(f"  {'CRL':<26}: {s['crl_line'] or '-'}")
        if not nt_nb:
            L.append(f"  {'Yolk sac':<26}: {d.get('yolk', 'Not Seen')}")
        L.append(f"  {'Cardiac pulsations':<26}: {twin_ft_cardiac(d)}")
        L.append(f"  {'Subchorionic collection':<26}: {s['subchorionic']}")
        L.append(f"  {'Placenta':<26}: {s['placenta']}")
        if nt_nb:
            L.append(f"  {'Choroid plexi / Falx':<26}: {s['choroid']}")
            L.append(f"  {'Nuchal translucency (NT)':<26}: {s['nt_val']}" + (f"  [{s['nt_pct']:g}%]" if s['nt_pct'] is not None else ""))
            L.append(f"  {'Nasal bone (NB)':<26}: {s['nb_val']}" + (f"  [{s['nb_pct']:g}%]" if s['nb_pct'] is not None else ""))
            if s["dv"]:
                L.append(f"  Ductus venosus: {s['dv']['text']}" + (f" ({s['dv']['status']})" if s['dv']['status'] else ""))
        L.append("")
    C = S["C"]
    L += ["COMMON FEATURES", f"  {'Corpus luteal cyst':<26}: {C['corpus']}", f"  {'Cervix':<26}: {C['cervix']}",
          f"  {'Internal Os':<26}: {C['internal_os']}", ""]
    if C["uta"]:
        u = C["uta"]
        L += ["PRE-ECLAMPSIA SCREEN:", f"  {u['text']}"]
        for lab, a, b, c_ in u["rows"]:
            L.append(f"  {lab:<24} PI {a}  RI {b}  S/D {c_}")
        L.append(f"  {'Mean Uterine Arteries':<24} PI {u['mean'][0] or '-'}  RI {u['mean'][1] or '-'}  S/D {u['mean'][2] or '-'}"
                 + (f"  [{u['pct']:g}%]" if u["pct"] is not None else ""))
        L.append("")
    L += ["OPINION:", twin_ft_impression(data, tw, nt_nb)]
    if twin_lowlying_advice(tw):
        L += ["", twin_lowlying_advice(tw)]
    return "\n".join(L)


# ============================================================
# ANOMALY SCAN TEMPLATE  (blue theme, 2 pages, own builder)
# ============================================================
ANOMALY = "ANOMALY"
# Anomaly Scan palettes: (title bar / headings / borders, label fill, thin lines, heading underline).
AN_PALETTES = {"Blue": ("1F4E79", "DDEBF7", "9DC3E6", "2E75B6"),
               "Pink": ("AD1457", "FCE4EC", "F48FB1", "EC407A")}
AN_NAVY, AN_LIGHT, AN_LINE, AN_MID = AN_PALETTES["Blue"]

def _an_set_palette(name):
    """Switch the colours every Anomaly helper below draws with (they read these globals when they run)."""
    global AN_NAVY, AN_LIGHT, AN_LINE, AN_MID
    AN_NAVY, AN_LIGHT, AN_LINE, AN_MID = AN_PALETTES.get(name, AN_PALETTES["Blue"])
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
        return f"{pos}, low-lying{lowlying_text(data)}" if pos else ""
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
    _tail = " and ".join(softmarker_phrases(data))
    _txt = f"Single live intra-uterine foetus in {pres}{',' if _is_lie_presentation(data) else ''} of Approx. {age} gestational age"
    return _txt + (f" with {_tail}." if _tail else ".")

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

def _ga_days(data):
    ga = _split_ga(data.get("mean_age"))
    if ga:
        return ga[0] * 7 + ga[1]
    return (_pf(data.get("ga_weeks_dop")) or 0) * 7

def _advice_for_ga(data):
    """Up to 34 weeks (34w0d) -> prompt consultation; beyond 34 weeks -> urgent consultation + delivery planning."""
    return AN_ADVICE_BEYOND_34 if _ga_days(data) > 34 * 7 else AN_ADVICE_UP_TO_34

_SM_COMMON = ("Correlation with first-trimester screening / maternal serum markers is advised, along with a detailed "
              "anomaly scan to look for other structural anomalies or soft markers if not already done. ")
SM_ADV_CPC = ("Adv- Isolated choroid plexus cysts are a benign finding and usually resolve by 28 weeks. " + _SM_COMMON +
              "Genetic counselling is advised if the screening is high risk or if any other abnormality is detected.")
SM_ADV_EIF = ("Adv- Isolated echogenic intracardiac focus is a common normal variant and does not indicate any cardiac "
              "dysfunction. " + _SM_COMMON +
              "Genetic counselling is advised if the screening is high risk or if any other abnormality is detected.")
SM_ADV_BOTH = ("Adv- The presence of more than one soft marker (choroid plexus cyst and echogenic intracardiac focus) "
               "slightly increases the likelihood of a chromosomal abnormality. " + _SM_COMMON +
               "Genetic counselling and consideration of non-invasive prenatal testing (NIPT) is advised.")
SM_DETAILED_SCAN = ", along with a detailed anomaly scan to look for other structural anomalies or soft markers if not already done"

def softmarker_advice(data, template_id):
    """Advice for a choroid plexus cyst / echogenic intracardiac focus / both. T16 under 24 weeks: starts with
    'Anomaly scan is advised.' and drops the 'detailed anomaly scan ... if not already done' phrase."""
    cp, ei = bool(cpc_phrase(data)), bool(eif_phrase(data))
    if not (cp or ei):
        return ""
    text = SM_ADV_BOTH if (cp and ei) else SM_ADV_CPC if cp else SM_ADV_EIF
    if template_id == "T16" and _ga_days(data) < 24 * 7:
        text = "Adv- Anomaly scan is advised. " + text[len("Adv- "):].replace(SM_DETAILED_SCAN, "")
    return text

def anomaly_advice(data):
    """Advice line under the opinion - only when any Doppler index is pathological."""
    if not any(r[4] for r in anomaly_doppler_rows(data)):
        return ""
    return _advice_for_ga(data)

AN_ADVICE_PSV = ("Adv- Evaluate for the cause of fetal anaemia: maternal blood group, Rh/atypical antibody screen (indirect Coombs) "
                 "and TORCH serology, Kleihauer-Betke test (fetomaternal haemorrhage), and haemoglobinopathy screening as "
                 "indicated. Repeat MCA PSV within 1 week, sooner if rising or if there is reduced fetal movement. Fetal blood "
                 "sampling and intrauterine transfusion may be considered at a tertiary centre as clinically indicated. "
                 "Interpret with caution beyond 35 weeks, as the false-positive rate rises.")

def _join_and(items):
    items = [i for i in items if i]
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " and " + items[-1]

def doppler_pi_findings(data, capitalise=True):
    """Pathological PI-type findings (UA / uterine Rt, Lt, Mean / MCA PI / CPR) as impression fragments,
    e.g. 'Increased Umbilical artery PI (1.45; 97.2th percentile)'. Percentile rules as in doppler_status()."""
    s = doppler_status(data)
    out = []
    def add(adj, label, value, pct):
        if not value:
            return
        a = adj if capitalise else adj.lower()
        pc = f"; {pct}th percentile" if pct not in (None, "") else ""
        out.append(f"{a} {label} ({value}{pc})")
    if s["ua"] == "Raised":
        add("Increased", "Umbilical artery PI", data.get("umbilical_pi"), data.get("umbilical_pct"))
    has_l = bool(data.get("uterine_left_pi"))
    has_r = bool(data.get("uterine_right_pi"))
    if s["uta_right"] == "Raised":
        add("Increased", "Right Uterine artery PI", data.get("uterine_right_pi"), data.get("uterine_right_pct"))
    if s["uta_left"] == "Raised":
        add("Increased", "Left Uterine artery PI", data.get("uterine_left_pi"), data.get("uterine_left_pct"))
    # Mean uterine PI: skipped when only one side was measured (the mean would just repeat that side)
    if s["uta"] == "Raised" and not (has_l != has_r):
        add("Increased", "Mean Uterine artery PI", data.get("uterine_pi"), data.get("uterine_pct"))
    if s["mca"] == "Reduced":
        add("Reduced", "MCA PI", data.get("mca_pi"), data.get("mca_pct"))
    if s["cpr"] == "Reduced":
        add("Reduced", "Cerebro-Placental Ratio", data.get("cpr"), data.get("cpr_pct"))
    return out

def doppler_psv_finding(data):
    """Impression fragment for a pathological MCA PSV (>= 1.5 MoM), else ''."""
    if doppler_status(data)["psv"] != "Raised":
        return ""
    return f"Raised MCA PSV at {data.get('mca_psv_mom')} MoM (\u22651.5 MoM), suspicious for fetal anaemia"

def doppler_advice_list(data):
    """Advisories for T3 / Color Doppler, each its own paragraph:
    MCA PSV >= 1.5 MoM -> fetal-anaemia advisory (first);
    any pathological PI finding (UA, uterine Rt/Lt/Mean, MCA PI, CPR) -> the GA-based advisory."""
    adv = []
    if doppler_psv_finding(data):
        adv.append(AN_ADVICE_PSV)
    s = doppler_status(data)
    if any(s[k] in ("Raised", "Reduced") for k in ("ua", "uta", "uta_left", "uta_right", "mca", "cpr")):
        adv.append(_advice_for_ga(data))
    return adv

def doppler_advice(data):
    """All advisories as one string (blank-line separated); empty when nothing is pathological."""
    return "\n\n".join(doppler_advice_list(data))

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

def _an_text_with_note(texts, normal_line, note):
    """Replace the 'normal' sentence by the soft-marker note when there is one."""
    return [note if (note and t == normal_line) else t for t in texts]

def anomaly_cranium_text(data):
    ph = cpc_phrase(data)
    return _an_text_with_note(AN_CRANIUM_TEXT, "No identifiable intracranial lesion seen.", f"Note is made of {ph}." if ph else "")

def anomaly_thorax_text(data):
    ph = eif_phrase(data)
    return _an_text_with_note(AN_THORAX_TEXT, "No evidence of SOL in the thorax.", f"Note is made of {ph}." if ph else "")

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

_AN_GAP_LINE = [6.5]          # blank-gap height; build_anomaly_docx_bytes lowers it when page 2 is unusually full

def _an_gap(doc, line=None):
    p = doc.add_paragraph()
    _an_flat(p, line=line or _AN_GAP_LINE[0])

def _an_lines(doc, texts):
    for t in texts:
        p = doc.add_paragraph()
        _an_flat(p, after=1, line=12.6)
        _an_run(p, t)

def _an_heading(doc, text, before=7):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()                       # border first, so python-docx keeps the schema order
    bd = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    for k, val in (("val", "single"), ("sz", "6"), ("space", "1"), ("color", AN_MID)):
        b.set(qn(f"w:{k}"), val)
    bd.append(b)
    pPr.append(bd)
    _an_flat(p, before=before, after=3)
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
    # Singleton Anomaly Scan rotates Blue / Pink with the report colour rotation (Blue if nothing was chosen)
    _an_set_palette(COLOR_THEMES[int(data.get("color_theme", ROTATION[0])) % len(COLOR_THEMES)][0])
    # Doppler advice AND soft-marker advice together make page 2 very full: smaller gaps / advice text keep it at 2 pages
    _two = sum(bool(x) for x in (anomaly_advice(data), softmarker_advice(data, ANOMALY), lowlying_advice(data))) >= 2
    _AN_GAP_LINE[0] = 2.5 if _two else 6.5
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
    _an_gap(doc, line=4)

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
    _an_heading(doc, "FOETAL SURVEY & ENVIRONMENT", before=5)
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

    _an_heading(doc, "CRANIUM", before=5)
    _an_measure_table(doc, AN_CRANIUM, data)
    _an_gap(doc, line=4)
    _an_lines(doc, anomaly_cranium_text(data))
    _an_heading(doc, "FOETAL SPINE", before=5)
    _an_lines(doc, AN_SPINE_TEXT)
    _an_heading(doc, "FACE", before=5)
    _an_lines(doc, AN_FACE_TEXT)
    _an_heading(doc, "ABDOMEN", before=5)
    _an_measure_table(doc, AN_ABDOMEN, data)
    _an_gap(doc, line=4)
    _an_lines(doc, AN_ABDOMEN_TEXT)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _an_flat(p, before=5)
    _an_run(p, "P.T.O", bold=True, color=AN_NAVY)

    # ---- page 2 top spacer: same 2.5 lines of 10 pt ----
    _an_spacer(doc, pt=10, page_break=True)
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=5)

    _an_heading(doc, "THORAX")
    _an_lines(doc, anomaly_thorax_text(data))
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
    sm_adv = softmarker_advice(data, ANOMALY)
    _adv_size, _adv_after = (8, 3) if _two else (9.5, 7)   # two advice paragraphs: smaller, so page 2 still holds the PNDT box
    if adv:                                   # pathological Doppler: advice goes on the next line, bold italics
        p = doc.add_paragraph()
        _an_flat(p, after=_adv_after)
        p.alignment = J
        _an_run(p, adv, bold=True, italic=True, size=_adv_size)
    if sm_adv:                                # soft-marker advice (choroid plexus cyst / EIF), bold italics
        p = doc.add_paragraph()
        _an_flat(p, after=_adv_after)
        p.alignment = J
        _an_run(p, sm_adv, bold=True, italic=True, size=_adv_size)
    if lowlying_advice(data):                 # low-lying placenta reaching / covering the Internal Os
        p = doc.add_paragraph()
        _an_flat(p, after=_adv_after)
        p.alignment = J
        _an_run(p, lowlying_advice(data), bold=True, italic=True, size=_adv_size)

    t = _an_table(doc, 1, [17.2])
    c = t.cell(0, 0)
    set_cell_background(c, AN_LIGHT)
    c.paragraphs[0].alignment = J
    _an_flat(c.paragraphs[0])
    _an_run(c.paragraphs[0], AN_PNDT, bold=True, italic=True, size=8.5, color=AN_NAVY)

    keep_row_together(t)
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
    if softmarker_advice(data, ANOMALY):
        L.append(f"  {softmarker_advice(data, ANOMALY)}")
    if lowlying_advice(data):
        L.append(f"  {lowlying_advice(data)}")
    for _n in softmarker_notes(data):
        L.append(f"  [Report body] {_n}")
    return "\n".join(L)


# ============================================================
# TWIN ANOMALY SCAN TEMPLATE  (same blue look as the singleton Anomaly Scan)
# Twin A (2 pages)  ->  Twin B (2 pages)  ->  common features (uterine arteries, cervix) + opinion + PNDT
# Everything the singleton Anomaly report carries is repeated for each twin. The ductus venosus is optional:
# it is printed (and judged) for a twin only when a DV PI was measured / typed for that twin.
# ============================================================
TWINS_ANOMALY = "TWINS_ANOMALY"
TWIN_AN_TEMPLATES = (TWINS_ANOMALY,)
REPORT_TITLES[TWINS_ANOMALY] = "USG OBS (TWIN PREGNANCY - ANOMALY SCAN)"
TWIN_AN_FULL = ("tcd", "ulna", "tibia")                      # size + weeks + days + percentile (besides BPD, HC, AC, FL, HL)
TWIN_AN_SIZE = ("lv", "cm", "nf", "radius", "fibula")        # size only
TWIN_AN_RATIOS = ("fl_ac", "fl_bpd", "fl_hc", "hc_ac")

# ---------------------------------------------------------------- text helpers
def _twan_ga_opinion(d):
    ga = _split_ga(d.get("mean_age"))
    return f"{ga[0]} weeks {ga[1]:02d} days \u00b1 2w" if ga else str(d.get("mean_age", "") or "____")

def twan_placenta(tw, data, X):
    """'Posterior, away from the Internal Os.' - Twin B follows Twin A's position when there is a single placenta."""
    return anomaly_placenta_text({**tw[X], "placenta_pos": twin_placenta_pos(tw, data, X)})

def twan_presentation(d):
    return twin_pres_one(d.get("presentation")).replace(" presentation", "")

def twan_cervix(data):
    raw = str(data.get("cervix", "") or "").strip()
    os_state = str(data.get("internal_os", "Closed") or "Closed").lower()
    try:
        float(raw.lower().replace("cm", "").replace("mm", "").strip())
        head = f"normal ({fmt_cm(raw)})"
    except ValueError:
        head = raw.lower() if raw else "normal"
    return f"Internal Os is {os_state}, Cervical length- {head}."

def _twan_rem(flags, ok="(WNL)."):
    return (", ".join(flags), True) if flags else (ok, False)

def twan_doppler_rows(d):
    """[(label, PI, percentile, remark, abnormal)] for ONE twin: UA always; MCA and ductus venosus only when measured."""
    s = doppler_status(d)
    rows = []
    v = d.get("umbilical_pi", "")
    rows.append(("UMBILICAL ARTERY PI", v, d.get("umbilical_pct", ""),
                 *(_twan_rem(["Raised PI"] if s["ua"] == "Raised" else []) if v else ("", False))))
    v = d.get("mca_pi", "")
    if v:
        rows.append(("MCA PI", v, d.get("mca_pct", ""), *_twan_rem(["Reduced PI"] if s["mca"] == "Reduced" else [])))
    v = str(d.get("dv_pi", "") or "").strip()
    if d.get("dv_on") and v:                                   # optional: carried into the report only if measured
        pct = _pf(d.get("dv_pct"))
        aw = str(d.get("dv_awave", "Positive") or "Positive")
        fl = (["Raised PI"] if (pct is not None and pct > 95) else []) + ([f"A-wave {aw.lower()}"] if aw != "Positive" else [])
        rows.append(("DUCTUS VENOSUS PI", v, d.get("dv_pct", ""), *_twan_rem(fl, "(WNL, Normal S, D & A Waves)")))
    return rows

def twan_uterine_rows(data):
    """Mean uterine artery PI (common to both twins) with the notching flag, as in the singleton Anomaly report."""
    s = doppler_status(data)
    notch = str(data.get("uterine_notching", "No notching"))
    fl = (["Raised PI"] if s["uta"] == "Raised" else []) + ([notch] if notch != "No notching" else [])
    v = data.get("uterine_pi", "")
    return [("MEAN UTERINE ARTERY PI", v, data.get("uterine_pct", ""), *(_twan_rem(fl) if v else ("", False)))]

def twan_any_doppler_abnormal(data, tw):
    return any(r[4] for X in "AB" for r in twan_doppler_rows(tw[X])) or any(r[4] for r in twan_uterine_rows(data))

def twan_softmarker_advice(tw):
    """Soft-marker advisories: a twin with both markers gets the 'more than one marker' text; CPC / EIF otherwise.
    Each distinct text is printed once, whichever twin it came from."""
    out = []
    for X in "AB":
        cp, ei = bool(cpc_phrase(tw[X])), bool(eif_phrase(tw[X]))
        if cp or ei:
            t = SM_ADV_BOTH if (cp and ei) else SM_ADV_CPC if cp else SM_ADV_EIF
            if t not in out:
                out.append(t)
    return out

def twan_advice_list(data, tw):
    adv = []
    if twan_any_doppler_abnormal(data, tw):
        adv.append(_advice_for_ga(max((tw["A"], tw["B"]), key=_ga_days)))     # GA of the more advanced twin
    adv += twan_softmarker_advice(tw)
    ll = twin_lowlying_advice(tw)
    if ll:
        adv.append(ll)
    return adv

def twan_opinion(data, tw):
    pre = "" if data.get("chorion_cert", "Likely") == "Established" else "Likely "
    s = (f"{pre}{twin_chorion_name(data)} twin live intra-uterine foetuses in {twin_pres_phrase(tw)}, "
         f"of Approx. {_twan_ga_opinion(tw['A'])} & {_twan_ga_opinion(tw['B'])} gestational age")
    s = s[0].upper() + s[1:]
    found = []
    for X in "AB":
        d = tw[X]
        found += [f"{ph} in Twin {X}" for ph in softmarker_phrases(d)]
        txt, abn = twin_liquor(d)
        if abn:
            kind = "anhydramnios" if txt.startswith("Anhydr") else "oligohydramnios" if txt.startswith("Reduced") else "polyhydramnios"
            found.append(f"anhydramnios in Twin {X}" if kind == "anhydramnios"
                         else f"{kind} in Twin {X} (largest pocket measures {d.get('pocket')} cm)")
    return s + (f", with {_join_and(found)}." if found else ".")

def twan_dating_rows(data, tw):
    def ga(s):
        g = _split_ga(s)
        return (f"{g[0]} Weeks", f"{g[1]} {'Day' if g[1] == 1 else 'Days'}") if g else ("-", "-")
    rows = []
    lmp = str(data.get("lmp", "") or "").strip()
    if lmp and lmp != "Not Sure":
        wk, dy = ga(data.get("ga_lmp"))
        rows.append(("By LMP", f"LMP: {_fmt_date(lmp)}", wk, dy, _fmt_date(data.get("edd_lmp", "")) or "-"))
    for X in "AB":
        wk, dy = ga(_tw_ga_text(tw[X]))
        rows.append((f"By USG (Twin {X})", "", wk, dy, _fmt_date(tw[X].get("edd_cua", "")) or "-"))
    return rows

# ---------------------------------------------------------------- form
def _twan_num(x):
    try:
        return float(str(x).replace(",", ".").replace("mm", "").strip())
    except (TypeError, ValueError):
        return None

def _twin_an_row(X, d, k, label, p, with_pct, lmp_not_sure):
    c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 3])
    c1.markdown(f"**{label}**")
    d[f"{p}_size"] = c2.text_input(f"{X} {p} size", key=k(f"{p}_size"), label_visibility="collapsed")
    d[f"{p}_wks"] = c3.text_input(f"{X} {p} wks", key=k(f"{p}_wks"), label_visibility="collapsed")
    d[f"{p}_dys"] = c4.text_input(f"{X} {p} dys", key=k(f"{p}_dys"), label_visibility="collapsed")
    if with_pct:
        pv = c5.text_input(f"{X} {p} pct", key=k(f"{p}_pct"), label_visibility="collapsed")
        d[f"{p}_pct"] = "" if lmp_not_sure else pv
        if pv and not lmp_not_sure:
            bar = generate_percentile_bar(pv)
            if bar: c5.image(bar, use_container_width=True)
    else:
        d[f"{p}_pct"] = ""

def _twin_an_hdr():
    h = st.columns([2, 1, 1, 1, 3])
    for c, t in zip(h, ("Parameter", "Size (mm)", "Wks", "Days", "Percentile")):
        c.caption(t)

def _twin_an_tab(X, lmp_not_sure, single_placenta):
    ss = st.session_state
    k = lambda n: f"sr_tw_{X}_{n}"
    d = {}
    b1, b2 = st.columns(2)
    if not ss.get(k("pres_set")):                                   # not chosen by the doctor yet: template default
        ss[k("pres")] = TWIN_DEFAULT_PRES.get(TWINS_ANOMALY, "VARIABLE")
    elif k("pres") not in ss and ss.get(k("pres_saved")) in TWIN_PRES_OPTS:
        ss[k("pres")] = ss[k("pres_saved")]                         # Streamlit forgets widgets not shown: restore
    d["presentation"] = b1.selectbox("Presentation", TWIN_PRES_OPTS, key=k("pres"),
                                     on_change=lambda _f=k("pres_set"): ss.__setitem__(_f, True))
    ss[k("pres_saved")] = d["presentation"]
    d["mean_age"] = b2.text_input("Mean age (GA by USG)", key=k("mean_age"))

    st.markdown("**Cranium**")
    _twin_an_hdr()
    _twin_an_row(X, d, k, "Biparietal Diameter", "bpd", True, lmp_not_sure)
    _twin_an_row(X, d, k, "Head Circumference", "hc", True, lmp_not_sure)
    _twin_an_row(X, d, k, "Cerebellum (TCD)", "tcd", True, lmp_not_sure)
    s1, s2, s3 = st.columns(3)
    d["lv_size"] = s1.text_input("Lateral ventricle (mm)", key=k("lv_size"))
    d["cm_size"] = s2.text_input("Cisterna magna (mm)", key=k("cm_size"))
    d["nf_size"] = s3.text_input("Nuchal fold (mm)", key=k("nf_size"))
    st.markdown("**Abdomen**")
    _twin_an_row(X, d, k, "Abdominal Circumference", "ac", True, lmp_not_sure)
    st.markdown("**Upper extremities**")
    _twin_an_row(X, d, k, "Humerus", "hl", True, lmp_not_sure)
    _twin_an_row(X, d, k, "Ulna", "ulna", True, lmp_not_sure)
    d["radius_size"] = st.text_input("Radius (mm)", key=k("radius_size"))
    st.markdown("**Lower extremities**")
    _twin_an_row(X, d, k, "Femur", "fl", True, lmp_not_sure)
    _twin_an_row(X, d, k, "Tibia", "tibia", True, lmp_not_sure)
    d["fibula_size"] = st.text_input("Fibula (mm)", key=k("fibula_size"))
    for p in ("radius", "fibula", "lv", "cm", "nf"):
        d.setdefault(f"{p}_wks", ""); d.setdefault(f"{p}_dys", ""); d.setdefault(f"{p}_pct", "")

    st.markdown("**Foetal indices** (worked out from the sizes - editable)")
    f_, h_, a_, b_ = (_twan_num(d.get(f"{p}_size")) for p in ("fl", "hc", "ac", "bpd"))
    calc = {}
    if f_ and h_: calc["fl_hc"] = f"{f_ / h_ * 100:.1f}%"
    if h_ and a_: calc["hc_ac"] = f"{h_ / a_:.2f}"
    if f_ and b_: calc["fl_bpd"] = f"{f_ / b_ * 100:.1f}%"
    if f_ and a_: calc["fl_ac"] = f"{f_ / a_ * 100:.1f}%"
    i1, i2 = st.columns(2)
    for (r, lab), col in zip((("fl_ac", "FL / AC"), ("fl_bpd", "FL / BPD"), ("fl_hc", "FL / HC"), ("hc_ac", "HC / AC")),
                             (i1, i2, i1, i2)):
        val = calc.get(r) or str(ss.get(k(f"{r}_sr"), "") or "")          # the machine's own ratio when sizes are missing
        if val and r != "hc_ac" and not val.endswith("%"):
            val += "%"
        d[r] = col.text_input(lab, value=val, key=k(f"{r}_{val}"))

    st.markdown("**Weight, heart rate, dating**")
    b1, b2, b3 = st.columns(3)
    d["fetal_weight"] = b1.text_input("Fetal weight (g)", key=k("weight"))
    d["efw_sd"] = b2.text_input("EFW \u00b1 (gms)  [blank = automatic 15%]", key=k("efw_sd"))
    d["fhr"] = b3.text_input("FHR (bpm)", key=k("fhr"))
    d["edd_cua"] = st.text_input("EDD (CUA)", key=k("edd_cua"))

    st.markdown("**Placenta, liquor**")
    b1, b2 = st.columns(2)
    if X == "B" and single_placenta:
        b1.caption("Single placenta - same as Twin A")
        d["placenta_pos"] = ""
    else:
        d["placenta_pos"] = b1.selectbox("Placenta position", PLACENTA_OPTS, key=k("plac_pos"))
    d["placenta_os"] = b2.selectbox("Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key=k("plac_os"))
    if d["placenta_os"] == "Low-lying":
        d["placenta_ll"] = lowlying_select(st, k("plac_ll"))
        if d["placenta_ll"] == PLAC_LL_OPTS[0]:
            d["placenta_dist"] = st.text_input("Distance from Internal Os (cm) [optional]", key=k("plac_dist"))
    elif d["placenta_os"] == "Placenta Praevia":
        d["praevia_grade"] = st.radio("Praevia grade", ["Grade I", "Grade II", "Grade III", "Grade IV"], horizontal=True, key=k("praevia"))
    d["pocket"] = st.text_input("Largest pocket (cm)", key=k("pocket"))
    _lt, _la = twin_liquor(d)
    if d["pocket"]:
        (st.warning if _la else st.caption)(_lt)

    st.markdown("**Soft markers**")
    d["cpc"] = st.checkbox("Choroid plexus cyst", key=k("cpc"))
    if d["cpc"]:
        k1, k2 = st.columns(2)
        d["cpc_right"] = k1.checkbox("Right", key=k("cpc_right"))
        if d["cpc_right"]:
            d["cpc_right_size"] = k1.text_input("Right cyst size (mm)", key=k("cpc_right_size"))
        d["cpc_left"] = k2.checkbox("Left", key=k("cpc_left"))
        if d["cpc_left"]:
            d["cpc_left_size"] = k2.text_input("Left cyst size (mm)", key=k("cpc_left_size"))
        if not (d["cpc_right"] or d["cpc_left"]):
            st.caption("Tick Right and/or Left.")
    d["eif"] = st.checkbox("Echogenic intracardiac focus (EIF)", key=k("eif"))
    if d["eif"]:
        e1, e2 = st.columns(2)
        d["eif_left"] = e1.checkbox("Left ventricle", key=k("eif_left"))
        d["eif_right"] = e2.checkbox("Right ventricle", key=k("eif_right"))
        if not (d["eif_left"] or d["eif_right"]):
            st.caption("Tick Left and/or Right ventricle.")

    st.markdown("**Doppler**")
    c1, c2, c3 = st.columns(3)
    d["umbilical_pi"] = c1.text_input("UA PI", key=k("ua_pi"))
    d["umbilical_ri"] = c2.text_input("UA RI", key=k("ua_ri"))
    d["umbilical_sd"] = c3.text_input("UA S/D", key=k("ua_sd"))
    _gd = _twin_ga_default(d["mean_age"])
    d["ga_weeks_dop"] = st.number_input("GA (weeks) for percentiles", min_value=10.0, max_value=42.0, value=_gd, step=0.1,
                                        key=k(f"ga_{_gd}"))
    _num = lambda s: float(str(s).replace(",", "."))
    if d["umbilical_pi"].strip():
        try:
            res = calculate_ua_pi(_num(d["umbilical_pi"]), d["ga_weeks_dop"])
            if res:
                d["umbilical_pct"] = str(res["percentile"])
                st.caption(f"Umbilical artery: {res['percentile']}th percentile")
                bar = generate_percentile_bar(res["percentile"])
                if bar: st.image(bar, width=220)
        except ValueError:
            st.caption("UA PI must be a number.")
    st.caption(f"Middle cerebral artery - Twin {X} (optional: the row appears only when a PI is entered)")
    c1, c2, c3 = st.columns(3)
    d["mca_pi"] = c1.text_input("MCA PI", key=k("mca_pi"))
    d["mca_ri"] = c2.text_input("MCA RI", key=k("mca_ri"))
    d["mca_sd"] = c3.text_input("MCA S/D", key=k("mca_sd"))
    try:
        if d["mca_pi"].strip():
            rm = calculate_mca_pi(_num(d["mca_pi"]), d["ga_weeks_dop"])
            if rm:
                d["mca_pct"] = str(rm["percentile"])
                st.caption(f"MCA: {rm['percentile']}th percentile")
                bar = generate_percentile_bar(rm["percentile"])
                if bar: st.image(bar, width=220)
            else:
                st.caption("No MCA percentile below 19 weeks.")
    except ValueError:
        st.caption("MCA PI must be a number.")
    d["dv_on"] = st.checkbox("Ductus venosus measured (optional - printed in the report only if a PI is entered)", key=k("dv_on"))
    if d["dv_on"]:
        c1, c2 = st.columns(2)
        d["dv_pi"] = c1.text_input("DV PI", key=k("dv_pi"))
        d["dv_awave"] = c2.selectbox("DV A-wave", ["Positive", "Absent", "Reversed"], key=k("dv_wave"))
        try:
            if d["dv_pi"].strip():
                rdv = calculate_dv_pi(_num(d["dv_pi"]), d["ga_weeks_dop"])
                if rdv:
                    d["dv_pct"] = str(rdv["percentile"])
                    st.caption(f"Ductus venosus: {rdv['percentile']}th percentile")
                    bar = generate_percentile_bar(rdv["percentile"])
                    if bar: st.image(bar, width=220)
                else:
                    st.caption("No DV percentile outside 20-40 weeks.")
        except ValueError:
            st.caption("DV PI must be a number.")
    return d

# ---------------------------------------------------------------- Word report
def _twan_page_top(doc, page_break=True):
    """Stationery gap at the top of a page (two 10 pt blank lines + one 5 pt, like the singleton report)."""
    _an_spacer(doc, pt=10, page_break=page_break)
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=10)
    _an_spacer(doc, pt=5)

def _twan_pto(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _an_flat(p, before=5)
    _an_run(p, "P.T.O", bold=True, color=AN_NAVY)

def _twan_banner(doc, X, d, cont=False):
    t = _an_table(doc, 1, [17.2])
    p = _an_cell(t.cell(0, 0), "", fill=AN_MID)
    _an_run(p, f"TWIN {X}", bold=True, size=11, color="FFFFFF")
    if X == "A":
        _an_run(p, "   (PRESENTING TWIN)", bold=True, size=8.5, color="FFFFFF")
    if cont:
        _an_run(p, "   (contd.)", bold=True, size=8.5, color="FFFFFF")
    _an_gap(doc, line=4)

def _twan_part1(doc, tw, data, X):
    """Banner, survey, cranium, spine, face."""
    d = tw[X]
    _twan_banner(doc, X, d)
    _an_heading(doc, "FOETAL SURVEY & ENVIRONMENT", before=2)
    liq, liq_abn = twin_liquor(d)
    survey = [("FHR", f"{str(d.get('fhr', '') or '').strip() or '____'} bpm - Regular", False),
              ("Presentation", twan_presentation(d), False),
              ("Placenta", twan_placenta(tw, data, X), False),
              ("Liquor", liq, liq_abn),
              ("E.F.W.", anomaly_efw(d), False)]
    t = _an_table(doc, len(survey), [3.6, 13.6])
    for i, (label, val, bold) in enumerate(survey):
        _an_label(t.rows[i].cells[0], label)
        _an_cell(t.rows[i].cells[1], val, bold=bold)
    _an_heading(doc, "CRANIUM", before=5)
    _an_measure_table(doc, AN_CRANIUM, d)
    _an_gap(doc, line=4)
    _an_lines(doc, anomaly_cranium_text(d))
    _an_heading(doc, "FOETAL SPINE", before=5)
    _an_lines(doc, AN_SPINE_TEXT)
    _an_heading(doc, "FACE", before=5)
    _an_lines(doc, AN_FACE_TEXT)

def _twan_part2(doc, tw, data, X):
    """Banner, abdomen, thorax, extremities, indices, Doppler."""
    d = tw[X]
    C = WD_ALIGN_PARAGRAPH.CENTER
    _twan_banner(doc, X, d, cont=True)
    _an_heading(doc, "ABDOMEN", before=2)
    _an_measure_table(doc, AN_ABDOMEN, d)
    _an_gap(doc, line=4)
    _an_lines(doc, AN_ABDOMEN_TEXT)
    _an_heading(doc, "THORAX")
    _an_lines(doc, anomaly_thorax_text(d))
    _an_heading(doc, "UPPER EXTREMITIES")
    _an_measure_table(doc, AN_UPPER, d)
    _an_heading(doc, "LOWER EXTREMITIES")
    _an_measure_table(doc, AN_LOWER, d)
    _an_gap(doc)
    _an_lines(doc, ["No obvious abnormality is noted in the extremities."])
    _an_heading(doc, "FOETAL INDICES")
    t = _an_table(doc, len(AN_INDICES), [6.0, 4.0, 7.2])
    for i, (label, kk, rng) in enumerate(AN_INDICES):
        c = t.rows[i].cells
        _an_label(c[0], label)
        _an_cell(c[1], str(d.get(kk, "") or ""), align=C)
        _an_cell(c[2], rng, align=C)
    _an_heading(doc, "DOPPLER INDICES")
    _twan_doppler_table(doc, twan_doppler_rows(d))

def _twan_doppler_table(doc, drows):
    C = WD_ALIGN_PARAGRAPH.CENTER
    t = _an_table(doc, len(drows), [5.6, 1.9, 5.2, 4.5])
    for i, (label, val, pct, remark, abn) in enumerate(drows):
        c = t.rows[i].cells
        _an_label(c[0], label)
        _an_cell(c[1], str(val or ""), align=C)
        _an_pct(c[2], pct)
        _an_cell(c[3], remark, bold=abn, align=C)

def build_twins_an_docx_bytes(data):
    tw = data["twins"]
    advs = twan_advice_list(data, tw)
    _tight = len(advs) >= 3
    _AN_GAP_LINE[0] = 2.5 if _tight else 6.5
    C, J = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.JUSTIFY
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(1.9)
    sec.header_distance = sec.footer_distance = Cm(0)
    st_ = doc.styles["Normal"]
    st_.font.name, st_.font.size = AN_FONT, Pt(9.5)
    st_.paragraph_format.space_after = st_.paragraph_format.space_before = Pt(0)

    # ---- page 1: title bar + patient details, dating, chorionicity ----
    _twan_page_top(doc, page_break=False)
    t = _an_table(doc, 3, [8.6, 8.6])
    title = t.cell(0, 0).merge(t.cell(0, 1))
    _an_cell(title, REPORT_TITLES[TWINS_ANOMALY], bold=True, color="FFFFFF", fill=AN_NAVY, align=C, size=12)
    age = str(data.get("patient_age", "") or "").strip()
    agesex = f"{age}Y/{data.get('p_sex', 'F')}" if age else ""
    _an_cell(t.cell(1, 0), f"NAME :-  {data.get('patient_name', '')}", bold=True)
    _an_cell(t.cell(1, 1), f"DATE :-  {data.get('report_date', '')}", bold=True)
    _an_cell(t.cell(2, 0), f"AGE/SEX :-  {agesex}", bold=True)
    _an_cell(t.cell(2, 1), f"REF. BY :-  {data.get('referred_by', '')}", bold=True)
    _an_gap(doc, line=4)

    rows = twan_dating_rows(data, tw)
    t = _an_table(doc, 1 + len(rows), [3.8, 4.0, 3.0, 3.0, 3.4])
    for ci, head in {0: "Dating", 1: "LMP", 4: "EDD"}.items():
        _an_cell(t.cell(0, ci), head, bold=True, color=AN_NAVY, fill=AN_LIGHT, align=C)
    _an_cell(t.cell(0, 2).merge(t.cell(0, 3)), "GA", bold=True, color=AN_NAVY, fill=AN_LIGHT, align=C)
    for ri, vals in enumerate(rows, start=1):
        for ci, txt in enumerate(vals):
            _an_cell(t.cell(ri, ci), txt, bold=(ci == 0), color=(AN_NAVY if ci == 0 else None), align=C)
    p = doc.add_paragraph()
    _an_flat(p, before=5, after=3, line=12.6)
    _an_run(p, "There are twin, live, normal intrauterine gestational pregnancies.")
    t = _an_table(doc, 1, [17.2])
    c = t.cell(0, 0)
    set_cell_background(c, AN_LIGHT)
    c.paragraphs[0].alignment = J
    _an_flat(c.paragraphs[0])
    _an_run(c.paragraphs[0], twin_chorion_sentence(data, tw), bold=True, italic=True, size=9, color=AN_NAVY)
    _an_gap(doc, line=5)

    # ---- Twin A (pages 1-2), Twin B (pages 3-4) ----
    _twan_part1(doc, tw, data, "A")
    _twan_pto(doc)
    _twan_page_top(doc)
    _twan_part2(doc, tw, data, "A")
    _twan_pto(doc)
    _twan_page_top(doc)
    _twan_part1(doc, tw, data, "B")
    _twan_pto(doc)
    _twan_page_top(doc)
    _twan_part2(doc, tw, data, "B")
    _twan_pto(doc)

    # ---- page 5: common features, opinion, PNDT ----
    _twan_page_top(doc)
    _an_heading(doc, "COMMON FEATURES (BOTH TWINS)", before=2)
    _twan_doppler_table(doc, twan_uterine_rows(data))
    _an_gap(doc, line=4)
    t = _an_table(doc, 1, [3.6, 13.6])
    _an_label(t.rows[0].cells[0], "Cervix")
    _an_cell(t.rows[0].cells[1], twan_cervix(data))
    for _ in range(2):
        _an_spacer(doc, pt=9.5)
    _an_heading(doc, "OPINION")
    p = doc.add_paragraph()
    _an_flat(p, after=3)
    _an_run(p, "Ultrasound Anomaly scan reveals:", bold=True, italic=True, size=7.5, color=AN_NAVY)
    p = doc.add_paragraph()
    _an_flat(p, after=7)
    _an_run(p, twan_opinion(data, tw), bold=True)
    _adv_size, _adv_after = (8, 3) if _tight else (9.5, 7)
    for txt in advs:
        p = doc.add_paragraph()
        _an_flat(p, after=_adv_after)
        p.alignment = J
        _an_run(p, txt, bold=True, italic=True, size=_adv_size)
    t = _an_table(doc, 1, [17.2])
    c = t.cell(0, 0)
    set_cell_background(c, AN_LIGHT)
    c.paragraphs[0].alignment = J
    _an_flat(c.paragraphs[0])
    _an_run(c.paragraphs[0], AN_PNDT, bold=True, italic=True, size=8.5, color=AN_NAVY)
    keep_row_together(t)
    finish_document(doc)
    normalize_table_props(doc)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def twins_an_preview(data):
    tw = data["twins"]
    L = [f"        {REPORT_TITLES[TWINS_ANOMALY]}", "",
         f"NAME :- {data.get('patient_name', '')}        DATE :- {data.get('report_date', '')}",
         f"AGE/SEX :- {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}        REF. BY :- {data.get('referred_by', '')}", "",
         f"  {'Dating':<18}{'LMP':<20}{'GA':<20}{'EDD'}"]
    for label, lmp, wk, dy, edd in twan_dating_rows(data, tw):
        L.append(f"  {label:<18}{lmp:<20}{(wk + ' ' + dy):<20}{edd}")
    L += ["", twin_chorion_sentence(data, tw), ""]

    def block(title, spec, d):
        L.append(title)
        for label, kk, has_pct in spec:
            row = f"  {label:<26}{_an_mm(d.get(kk + '_size', '')):<8}{_an_weeks(d.get(kk + '_wks', '')):<10}{_an_days(d.get(kk + '_dys', '')):<9}"
            if has_pct and d.get(kk + "_pct"):
                row += f"{d.get(kk + '_pct')}%"
            L.append(row.rstrip())
        L.append("")
    for X in "AB":
        d = tw[X]
        L += [f"TWIN {X}" + ("  (PRESENTING TWIN)" if X == "A" else ""),
              "FOETAL SURVEY & ENVIRONMENT",
              f"  FHR           : {str(d.get('fhr', '') or '').strip() or '____'} bpm - Regular",
              f"  Presentation  : {twan_presentation(d)}",
              f"  Placenta      : {twan_placenta(tw, data, X)}",
              f"  Liquor        : {twin_liquor(d)[0]}",
              f"  E.F.W.        : {anomaly_efw(d)}", ""]
        block("CRANIUM", AN_CRANIUM, d); block("ABDOMEN", AN_ABDOMEN, d)
        block("UPPER EXTREMITIES", AN_UPPER, d); block("LOWER EXTREMITIES", AN_LOWER, d)
        L.append("FOETAL INDICES")
        for label, kk, rng in AN_INDICES:
            L.append(f"  {label:<18}{str(d.get(kk, '') or ''):<10}{rng}")
        L += ["", "DOPPLER INDICES"]
        for label, val, pct, remark, _ in twan_doppler_rows(d):
            L.append(f"  {label:<26}{str(val or ''):<7}{(str(pct) + '%') if pct else '':<9}{remark}")
        for _n in softmarker_notes(d):
            L.append(f"  [Report body] {_n}")
        L.append("")
    L.append("COMMON FEATURES (BOTH TWINS)")
    for label, val, pct, remark, _ in twan_uterine_rows(data):
        L.append(f"  {label:<26}{str(val or ''):<7}{(str(pct) + '%') if pct else '':<9}{remark}")
    L.append(f"  Cervix: {twan_cervix(data)}")
    L += ["", "OPINION", "  Ultrasound Anomaly scan reveals:", f"  {twan_opinion(data, tw)}"]
    for a in twan_advice_list(data, tw):
        L.append(f"  {a}")
    return "\n".join(L)


# ============================================================
# LIVE PREVIEW RENDERER
# ============================================================
def render_ob_preview(data, template_id):
    if is_twin_template(template_id):
        return twins_preview(data, template_id)
    if template_id == ANOMALY:
        return render_anomaly_preview(data)
    if template_id == IUD_OBS:
        return iud_preview(data)
    lines = []
    lines.append(f"        {REPORT_TITLES.get(template_id, 'USG OBS')}")
    lines.append(f"NAME :-  {data.get('patient_name', '')}        DATE :-  {data.get('report_date', '')}")
    lines.append(f"AGE/SEX :-  {data.get('patient_age', '')}Y/{data.get('p_sex', 'F')}        REF. BY :-  {data.get('referred_by', '')}")
    lines.append("")

    if template_id == "EARLY_OBS":
        lines.extend(early_preview(data))

    elif template_id == "NT_NB":
        lines.extend(nt_nb_preview(data))

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
            plac_str = f"is {_pos_str} and Low-lying" + lowlying_text(data)
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
        for _n in softmarker_notes(data):
            lines.append(_n)

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

    lines.append(""); lines.append("OPINION:" if template_id in ("NT_NB", "EARLY_OBS") else "IMPRESSION:")
    mean_age = data.get('mean_age', '')
    
    imp_map = {
        "Blighted": f"INTRA UTERINE EMPTY GESTATIONAL SAC OF {data.get('ga_crl', '')} +/-1w WITH NO APPRECIABLE FETAL POLE OR YOLK SAC - LIKELY EARLY PREGNANCY FAILURE (BLIGHTED OVUM).",
        "NotYet": f"Intra uterine pregnancy of {data.get('ga_crl', '')} +/-1w with not yet appreciable cardiac activity. Adv- Review scan after 2 weeks.",
        "Present": f"Intra uterine live pregnancy of {data.get('ga_crl', '')} +/-1w with good cardiac activity.",
        "NT_NB": nt_nb_opinion(data)
    }

    if template_id == "EARLY_OBS":
        imp_text = early_opinion(data)
    elif template_id == "NT_NB":
        imp_text = imp_map["NT_NB"]
    else:
        pres_text = presentation_text(data)
            
        imp_text = f"Ultrasound obstetric reveals single live intrauterine fetus in {pres_text}{',' if _is_lie_presentation(data) else ''} of approx. {mean_age} gestation age"
        _psv_f = doppler_psv_finding(data) if template_id in ("T3", COLOR_DOPPLER) else ""
        _pi_f = doppler_pi_findings(data, capitalise=(template_id == "T3")) if template_id in ("T3", COLOR_DOPPLER) else []
        _inline = softmarker_phrases(data) + (_pi_f if template_id == "T3" else []) + ([_psv_f] if _psv_f else [])
        imp_text += impression_with_clause(_inline, data)
        imp_text += "."
        
    lines.append(f"  - {imp_text}")
    if template_id == COLOR_DOPPLER and doppler_status(data)["all_normal"]:
        lines.append("")
        lines.append(f"  - {NORMAL_DOPPLER_LINE}")
    if template_id == COLOR_DOPPLER and _pi_f:
        lines.append("")
        lines.append(f"  - Doppler study reveals {_join_and(_pi_f)}.")
    _sf = smaller_fetus_remark(data) if template_id not in ("EARLY_OBS", "NT_NB") else ""
    if _sf:
        lines.append(""); lines.append(f"  - {_sf}")
    for _a in ((doppler_advice_list(data) if template_id in ("T3", COLOR_DOPPLER) else [])
               + ([softmarker_advice(data, template_id)] if template_id in ("T16", "T3", COLOR_DOPPLER)
                  and softmarker_advice(data, template_id) else [])
               + ([lowlying_advice(data)] if lowlying_advice(data) and template_id != "EARLY_OBS" else [])):
        lines.append("")
        lines.append(f"  {_a}")

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

# Fields the OCR of the ultrasound images can fill, and the widget keys that hold them (all templates).
OCR_FORM_KEYS = ["placenta_pos_widget", "placenta_os_widget", "placenta_ll_widget", "nt_plac_ll", "nt_cervix",
                 "iud_plac_pos", "iud_plac_os", "iud_plac_ll", "iud_afi_method", "iud_afi_total", "iud_afi_pocket",
                 "iud_cervix", "sr_tw_c_cervix", "early_plac_pos", "early_cl", "early_cl_side", "early_cl_size"]

# Everything the doctor enters for ONE exam that is not stored under a versioned "sr_..._{v}" key.
# Cleared when a new patient starts, so presentation / lie / placenta can never carry over to the next patient.
EXAM_WIDGET_PREFIXES = ("sr_ocr_", "presentation_widget", "auto_fallback_")
EXAM_WIDGET_EXACT = set(OCR_FORM_KEYS) | {"pres_user_set", "pres_saved", "lie_widget", "lie_widget_disabled", "head_side_widget",
                                          "head_level_widget", "_ocr_snap", "_ocr_touched"}
def _clear_exam_widgets():
    for _k in list(st.session_state.keys()):
        if _k in EXAM_WIDGET_EXACT or _k.startswith(EXAM_WIDGET_PREFIXES):
            del st.session_state[_k]

def _mark_pres_set():
    st.session_state["pres_user_set"] = True      # the doctor has chosen the presentation: never change it automatically

# Placenta widgets that the OCR may fill. If the doctor changes one by hand, the OCR must not overwrite it afterwards.
PLAC_WATCH_KEYS = ["sr_ocr_plac_pos", "sr_ocr_plac_os", "placenta_pos_widget", "placenta_os_widget", "iud_plac_pos", "iud_plac_os", "early_plac_pos"]

if os.path.exists(json_path):
    mtime = None
    try:
        mtime = os.path.getmtime(json_path)
        if st.session_state.get("last_json_mtime") != mtime and st.session_state.get("json_bad_mtime") != mtime:
            with open(json_path, 'r') as f:
                new_data = json.load(f)
            _prev = st.session_state.get("sr_data") or {}
            _same_study = bool(new_data.get("_study")) and new_data.get("_study") == _prev.get("_study")
            # Same exam and same SR, only new OCR text from an image: top the values up WITHOUT starting a new form version
            # (a new version would reset everything the SR filled in and that has been corrected by hand since).
            _ocr_only = (_same_study and st.session_state.get("last_json_mtime") is not None
                         and str(new_data.get("_sr_id", "")) == str(_prev.get("_sr_id", "")))
            if _ocr_only:
                st.session_state["sr_data"] = new_data
                st.session_state["last_json_mtime"] = mtime
                st.toast("🔍 Read text from a new ultrasound image (OCR)")
            elif not _same_study and _prev.get("_study"):  # a different exam replaces a loaded one: forget the previous patient's entries
                _clear_exam_widgets()
                # (an EMPTY form is NOT wiped: whatever the doctor entered before the images / SR arrived belongs to this exam)
        if (st.session_state.get("last_json_mtime") != mtime and st.session_state.get("json_bad_mtime") != mtime):
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

def apply_ocr_values(sr):
    """Put what the OCR read from the images into the form widgets (placenta, AFI / largest pocket, scar, cervix).
    A field is (re)written when its OCR reading is new or changed - so a value corrected by hand stays until a newer reading
    arrives - and also when the widget has lost its value (Streamlit forgets a widget's value while it is not on screen,
    e.g. after switching templates)."""
    ocr = {k[4:]: val for k, val in (sr or {}).items() if k.startswith("ocr_") and k != "ocr_src" and isinstance(val, str)}
    ss = st.session_state
    # which placenta widgets has the doctor changed by hand since the last run? (value differs from what we left there)
    touched = ss.setdefault("_ocr_touched", set())
    for _k, _val in ss.get("_ocr_snap", {}).items():
        if _k in ss and ss[_k] != _val:
            touched.add(_k)
    def _snapshot():
        ss["_ocr_snap"] = {_k: ss[_k] for _k in PLAC_WATCH_KEYS if _k in ss}
    if not ocr:
        _snapshot()
        return
    done = ss.get("sr_ocr_applied", {})
    def put(field, keys, val):
        fresh = done.get(field) != ocr.get(field)
        for _k in keys:
            if _k in touched and _k in ss:            # corrected by hand: the doctor's choice wins
                continue
            if fresh or _k not in ss:
                ss[_k] = val
    if ocr.get("placenta_pos") in PLACENTA_OPTS:
        put("placenta_pos", ["sr_ocr_plac_pos", "placenta_pos_widget", "iud_plac_pos", "early_plac_pos"], ocr["placenta_pos"])
    if ocr.get("placenta_os") in ("away from the Os", "Low-lying", "Placenta Praevia"):
        put("placenta_os", ["sr_ocr_plac_os", "placenta_os_widget", "iud_plac_os"], ocr["placenta_os"])
    if ocr.get("placenta_ll") in PLAC_LL_OPTS:
        put("placenta_ll", ["placenta_ll_widget", "nt_plac_ll", "iud_plac_ll"], ocr["placenta_ll"])
    if "cervix" in ocr:
        put("cervix", ["sr_ocr_cervix", "sr_ocr_cervix_early", "nt_cervix", "iud_cervix", "sr_tw_c_cervix"], ocr["cervix"])
    if "cl_cyst" in ocr:                      # Early OBS: machine annotation "CL CYST RT OVARY" / "CL CYST LT OVARY"
        _cl_side, _cl_size = parse_cl_cyst(ocr["cl_cyst"])
        if _cl_side:
            put("cl_cyst", ["early_cl"], True)
            put("cl_cyst", ["early_cl_side"], _cl_side)
            if _cl_size:
                put("cl_cyst", ["early_cl_size"], _cl_size)
    if "scar" in ocr:
        put("scar", ["sr_ocr_scar_size"], ocr["scar"])
        put("scar", ["sr_ocr_measure_scar"], True)
    if "afi_pocket" in ocr:
        put("afi_pocket", ["sr_ocr_afi_pocket", "iud_afi_pocket"], ocr["afi_pocket"])
    if "afi_total" in ocr:
        put("afi_total", ["sr_ocr_afi_total", "iud_afi_total"], ocr["afi_total"])
    if "afi_pocket" in ocr or "afi_total" in ocr:
        method = "Largest Pocket" if ocr.get("afi_pocket") else "Total 4 Quadrants"
        fresh = done.get("afi_pocket") != ocr.get("afi_pocket") or done.get("afi_total") != ocr.get("afi_total")
        for _k in ("sr_ocr_afi_method", "iud_afi_method"):
            if fresh or _k not in ss:
                ss[_k] = method
    ss["sr_ocr_applied"] = dict(ocr)
    _snapshot()

# Seed the placenta widgets BEFORE the OCR step so a manual change is always noticed (and the defaults are known)
for _k, _dv in (("sr_ocr_plac_pos", PLACENTA_OPTS[0]), ("iud_plac_pos", PLACENTA_OPTS[0]),
                ("placenta_pos_widget", "Anterior"), ("early_plac_pos", "-"),
                ("sr_ocr_plac_os", "away from the Os"), ("iud_plac_os", "away from the Os"),
                ("placenta_os_widget", "away from the Os")):
    if _k not in st.session_state:
        st.session_state[_k] = _dv

v = st.session_state.get("dicom_version", 0)
sr_data = st.session_state.get("sr_data", {})
apply_ocr_values(sr_data)

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

if "p_date_input" not in st.session_state:
    st.session_state["p_date_input"] = datetime.now().strftime("%d-%b-%y")
if "p_ref_input" not in st.session_state:
    st.session_state["p_ref_input"] = ""
_sr_demo_v = st.session_state.get("dicom_version", 0)
if sr_data and st.session_state.get("demo_applied_v") != _sr_demo_v:     # scan date / referring doctor from the SR, once per load
    if sr_data.get("study_date"): st.session_state["p_date_input"] = sr_data["study_date"]
    if sr_data.get("referred_by"): st.session_state["p_ref_input"] = sr_data["referred_by"]
    st.session_state["demo_applied_v"] = _sr_demo_v

c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 2, 3])
with c1: p_name = st.text_input("Name", key="p_name_input")
with c2: p_age = st.text_input("Age", key="p_age_input")
with c3: p_sex = st.radio("Sex", ["F", "M"], horizontal=True)
with c4: p_date = st.text_input("Date", key="p_date_input")
with c5: p_ref = st.text_input("Referred by", key="p_ref_input")

data = {"patient_name": p_name, "patient_age": p_age, "report_date": p_date, "referred_by": p_ref, "p_sex": p_sex}

# --- NEW PATIENT BUTTON ---
if st.button("🚀 Start New Patient / Clear Form", type="primary"):
    # Clear the DICOM data
    st.session_state["sr_data"] = {}
    st.session_state["dicom_version"] += 1
    # Clear the input fields by deleting their session state keys
    for key in list(st.session_state.keys()):
        if key.startswith("sr_") or key.startswith("tw_") or key in ["p_name_input", "p_age_input"]:
            del st.session_state[key]
    _clear_exam_widgets()          # presentation, lie, placenta, head position, template fallback choice ...
    st.rerun()
# ---------------------------

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    # Separate template menus per pregnancy type. Twins / IUD templates are added one at a time:
    # entries in COMING_SOON are shown in the list but have no report yet.
    TEMPLATE_MENU = {
        "Singleton": [("EARLY_OBS", "EARLY_OBS"), ("NT_NB", "NT_NB"), ("OBS_AUTO", "OBS (auto: T16 / T3 from SR)"),
                      (COLOR_DOPPLER, COLOR_DOPPLER), (ANOMALY, ANOMALY)],
        "Twins": [("TWINS_EARLY", "Early OBS"), ("TWINS_NT_NB", "NT / NB"),
                  ("TWINS_AUTO", "OBS (auto: T16 / T3 from SR)"), (TWINS_IUD, "IUD (fetal demise)"), ("TWINS_DOPPLER", "Doppler"), ("TWINS_ANOMALY", "Anomaly")],
        "IUD": [(IUD_OBS, "OBS (2nd / 3rd trimester)"), ("IUD_EARLY", "Early OBS (coming soon)")],
    }
    COMING_SOON = {"IUD_EARLY"}
    preg_type = st.radio("Pregnancy", list(TEMPLATE_MENU), horizontal=True, key="preg_type")
    _menu = TEMPLATE_MENU[preg_type]
    _tkey = f"template_select_{preg_type}"
    if st.session_state.get(_tkey) not in [t for t, _ in _menu]:     # a choice saved by an older version (e.g. IUD_T16) no longer exists
        st.session_state.pop(_tkey, None)
    template_id = st.selectbox("Template", [t for t, _ in _menu], format_func=dict(_menu).get, key=_tkey)
    if template_id in COMING_SOON:
        st.info(f"The {preg_type} {dict(_menu)[template_id].replace(' (coming soon)', '')} template has not been built yet.")
        st.stop()
    
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

    # --- AUTO TEMPLATE: the SR gestational age picks T16 (< 25 weeks) or T3 (>= 25 weeks) ---
    if template_id in ("OBS_AUTO", "TWINS_AUTO"):
        _is_tw = template_id == "TWINS_AUTO"
        _auto = auto_obs_template(sr_data, twins=_is_tw)
        if _auto is None:
            st.warning("No gestational age in the SR yet - load the SR file to auto-select T16 / T3, or pick one for now.")
            _pick = st.radio("Template until SR arrives", ["T16", "T3"], horizontal=True, key=f"auto_fallback_{_is_tw}")
            _auto = ("TWINS_" if _is_tw else "") + _pick
        else:
            _d = sr_ga_days(sr_data)
            st.success(f"SR gestational age {_d // 7}w{_d % 7}d -> template {_auto.replace('TWINS_', '')} selected automatically")
        template_id = _auto
    
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
        if 'nb_len' in sr_data: st.session_state[f"sr_nb_len_{v}"] = sr_data['nb_len']
        if 'nb_pct' in sr_data: st.session_state[f"sr_nb_pct_{v}"] = sr_data['nb_pct']
        if 'crl_wks' in sr_data:
            st.session_state[f"sr_ga_crl_{v}"] = f"{sr_data['crl_wks']}wks {int(sr_data.get('crl_dys', 0)):02d}days"
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
    _ocr_src = (sr_data or {}).get("ocr_src")
    if _ocr_src:
        _ocr_vals = {k[4:]: val for k, val in sr_data.items() if k.startswith("ocr_") and k != "ocr_src"}
        with st.expander("🔍 Filled from the ultrasound images (OCR) - please check: " + ", ".join(sorted(_ocr_vals)), expanded=False):
            for _f, _t in _ocr_src.items():
                st.markdown(f"**{_f.replace('_', ' ')}** = {_ocr_vals.get(_f, '')}  &nbsp; ← read as *\"{_t}\"*")
            st.caption("These are ordinary editable fields. A field is overwritten only when a newer image gives a new reading.")
    lmp_not_sure = st.checkbox("LMP Not Sure / Not Known")
    if lmp_not_sure:
        data["lmp"] = "Not Sure"
        data["ga_lmp"] = ""
    else:
        data["lmp"] = st.text_input("LMP", value=st.session_state.get(f"sr_lmp_{v}", ""))
        data["ga_lmp"] = st.text_input("GA by LMP", value=st.session_state.get(f"sr_ga_lmp_{v}", ""), key=f"ga_lmp_input_{v}")

    import contextlib
    with (st.expander("🔗 DICOM SR Data (Machine Generated)", expanded=False) if not is_twin_template(template_id) else contextlib.nullcontext()):
        if not is_twin_template(template_id):
            st.markdown('<div class="dicom-box">Verify or override data received from the ultrasound machine.</div>', unsafe_allow_html=True)
        
        if is_twin_template(template_id):
            st.markdown("### 👶 Twin Pregnancy")
            twins_form(data, sr_data, v, lmp_not_sure, template_id)
        if template_id != "EARLY_OBS" and not is_twin_template(template_id):
            # Timeline Header
            st.markdown(f"**LMP:** {st.session_state.get(f'sr_lmp_{v}', '')} &nbsp;|&nbsp; "
                        f"**GA(LMP):** {st.session_state.get(f'sr_ga_lmp_{v}', '')} &nbsp;|&nbsp; "
                        f"**AUA:** {st.session_state.get(f'sr_mean_age_{v}', '')} &nbsp;|&nbsp; "
                        f"**EDD(LMP):** {st.session_state.get(f'sr_edd_lmp_{v}', '')}")
            st.markdown("---")

        if template_id == "EARLY_OBS":
            b1, b2 = st.columns(2)
            data["crl"] = b1.text_input("CRL (mm)", value=st.session_state.get(f"sr_crl_{v}", ""), key=f"sr_crl_{v}")
            data["msd"] = b2.text_input("G-Sac / MSD (mm) - leave blank if not measured (prints 'Seen')",
                                        value=st.session_state.get(f"sr_msd_{v}", ""), key=f"sr_msd_{v}")
            data["ga_crl"] = st.text_input("GA by CRL", value=st.session_state.get(f"sr_ga_crl_{v}", ""), key=f"sr_ga_crl_{v}")
            st.caption("Uterine Arteries (pre-eclampsia screen) - leave blank if not done")
            c1, c2, c3 = st.columns(3)
            data["uterine_right_pi"] = c1.text_input("Right PI", key=f"sr_uterine_right_pi_{v}")
            data["uterine_right_ri"] = c2.text_input("Right RI", key=f"sr_uterine_right_ri_{v}")
            data["uterine_right_sd"] = c3.text_input("Right S/D", key=f"sr_uterine_right_sd_{v}")
            c1, c2, c3 = st.columns(3)
            data["uterine_left_pi"] = c1.text_input("Left PI", key=f"sr_uterine_left_pi_{v}")
            data["uterine_left_ri"] = c2.text_input("Left RI", key=f"sr_uterine_left_ri_{v}")
            data["uterine_left_sd"] = c3.text_input("Left S/D", key=f"sr_uterine_left_sd_{v}")

            def _early_mean_of(*vals):
                nums = []
                for x in vals:
                    try: nums.append(float(x))
                    except (TypeError, ValueError): pass
                return f"{sum(nums)/len(nums):.2f}" if nums else ""
            _mpi = _early_mean_of(data["uterine_left_pi"], data["uterine_right_pi"]) or sr_data.get("uterine_pi", "")
            _mri = _early_mean_of(data["uterine_left_ri"], data["uterine_right_ri"]) or sr_data.get("uterine_ri", "")
            _msd = _early_mean_of(data["uterine_left_sd"], data["uterine_right_sd"]) or sr_data.get("uterine_sd", "")
            c1, c2, c3 = st.columns(3)
            data["uterine_pi"] = c1.text_input("Mean Uterine PI", value=_mpi, key=f"sr_uterine_pi_mean_{v}_{_mpi}")
            data["uterine_ri"] = c2.text_input("Mean Uterine RI", value=_mri, key=f"sr_uterine_ri_mean_{v}_{_mri}")
            data["uterine_sd"] = c3.text_input("Mean Uterine S/D", value=_msd, key=f"sr_uterine_sd_mean_{v}_{_msd}")

            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
        elif template_id == "NT_NB":
            b1, b2 = st.columns(2)
            data["crl"] = b1.text_input("CRL (mm)", value=st.session_state.get(f"sr_crl_{v}", ""), key=f"sr_crl_{v}")
            data["nt"] = b2.text_input("NT (mm)", value=st.session_state.get(f"sr_nt_{v}", ""), key=f"sr_nt_{v}")
            
            b1, b2, b3 = st.columns(3)
            data["nasal_bone"] = b1.text_input("Nasal Bone (Present / Absent)", value=st.session_state.get(f"sr_nasal_bone_{v}", ""), key=f"sr_nasal_bone_{v}")
            data["nb_len"] = b2.text_input("Nasal Bone length (mm)", value=st.session_state.get(f"sr_nb_len_{v}", ""), key=f"sr_nb_len_{v}")
            data["nb_pct"] = b3.text_input("Nasal Bone percentile (machine)", value=st.session_state.get(f"sr_nb_pct_{v}", ""), key=f"sr_nb_pct_{v}")

            b1, b2 = st.columns(2)
            data["fhr_val"] = b1.text_input("FHR (bpm)", value=st.session_state.get(f"sr_fhr_{v}", ""), key=f"sr_fhr_{v}")
            data["ga_crl"] = b2.text_input("GA by CRL", value=st.session_state.get(f"sr_ga_crl_{v}", ""), key=f"sr_ga_crl_{v}")
            st.caption("Uterine Arteries (pre-eclampsia screen)")
            c1, c2, c3 = st.columns(3)
            data["uterine_right_pi"] = c1.text_input("Right PI", key=f"sr_uterine_right_pi_{v}")
            data["uterine_right_ri"] = c2.text_input("Right RI", key=f"sr_uterine_right_ri_{v}")
            data["uterine_right_sd"] = c3.text_input("Right S/D", key=f"sr_uterine_right_sd_{v}")
            c1, c2, c3 = st.columns(3)
            data["uterine_left_pi"] = c1.text_input("Left PI", key=f"sr_uterine_left_pi_{v}")
            data["uterine_left_ri"] = c2.text_input("Left RI", key=f"sr_uterine_left_ri_{v}")
            data["uterine_left_sd"] = c3.text_input("Left S/D", key=f"sr_uterine_left_sd_{v}")

            def _nt_mean_of(*vals):
                nums = []
                for x in vals:
                    try: nums.append(float(x))
                    except (TypeError, ValueError): pass
                return f"{sum(nums)/len(nums):.2f}" if nums else ""
            _mpi = _nt_mean_of(data["uterine_left_pi"], data["uterine_right_pi"]) or sr_data.get("uterine_pi", "")
            _mri = _nt_mean_of(data["uterine_left_ri"], data["uterine_right_ri"]) or sr_data.get("uterine_ri", "")
            _msd = _nt_mean_of(data["uterine_left_sd"], data["uterine_right_sd"]) or sr_data.get("uterine_sd", "")
            c1, c2, c3 = st.columns(3)
            data["uterine_pi"] = c1.text_input("Mean Uterine PI", value=_mpi, key=f"sr_uterine_pi_mean_{v}_{_mpi}")
            data["uterine_ri"] = c2.text_input("Mean Uterine RI", value=_mri, key=f"sr_uterine_ri_mean_{v}_{_mri}")
            data["uterine_sd"] = c3.text_input("Mean Uterine S/D", value=_msd, key=f"sr_uterine_sd_mean_{v}_{_msd}")

            data["edd_lmp"] = st.text_input("EDD (LMP)", value=st.session_state.get(f"sr_edd_lmp_{v}", ""), key=f"sr_edd_lmp_{v}")
            data["edd_cua"] = st.text_input("EDD (CUA)", value=st.session_state.get(f"sr_edd_cua_{v}", ""), key=f"sr_edd_cua_{v}")
        elif not is_twin_template(template_id):
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
            b1, b2, b3 = st.columns(3)
            data["cardiac_status"] = b1.selectbox("Cardiac pulsations", ["Present", "Not yet appreciable", "Blighted Ovum"], key="early_cardiac")
            if data["cardiac_status"] == "Present":
                data["fhr_val"] = b2.text_input("FHR (bpm)", value=st.session_state.get(f"sr_fhr_{v}", ""), key=f"sr_fhr_{v}")
            data["yolk"] = b3.selectbox("Yolk sac", ["Not seen", "Seen"], key="early_yolk")
            b1, b2, b3 = st.columns(3)
            data["decidual"] = b1.selectbox("Decidual reaction", ["Present", "Absent"], key="early_decidual")
            st.session_state.setdefault("sr_ocr_cervix_early", "Normal")
            data["cervix"] = b2.text_input("Cervix length (cm)", key="sr_ocr_cervix_early")
            data["internal_os"] = b3.selectbox("Internal Os", ["Closed", "Open"], key="early_os")
            st.caption(f"Cervix prints as: {early_cervix_text(data['cervix'])}   (type the length, e.g. 4.4 - below 2.5 cm prints as Short)")
            b1, b2 = st.columns(2)
            data["subchorionic_present"] = b1.checkbox("Subchorionic collection present", key="early_sch")
            if data["subchorionic_present"]:
                data["subchorionic_grade"] = b1.radio("Grade", ["Grade-I", "Grade-I/II", "Grade-II", "Grade-II/III"], horizontal=True, key="early_sch_g")
            data["corpus_present"] = b2.checkbox("Corpus luteal cyst present (OCR: CL CYST RT / LT OVARY)", key="early_cl")
            if data["corpus_present"]:
                d1, d2 = b2.columns(2)
                data["corpus_side"] = d1.selectbox("Side", ["Right", "Left", "Both"], key="early_cl_side")
                if data["corpus_side"] != "Both":
                    data["corpus_size"] = d2.text_input("Size (mm)", key="early_cl_size")
            b1, b2 = st.columns(2)
            _early_pp = b1.selectbox("Placenta", ["-"] + PLACENTA_OPTS,
                                     format_func=lambda o: "Not formed yet (shows  -  )" if o == "-" else o, key="early_plac_pos")
            data["placenta_pos"] = "" if _early_pp == "-" else _early_pp
            if _early_pp != "-":
                data["placenta_os"] = b2.selectbox("Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key="placenta_os_widget")
                if data["placenta_os"] == "Low-lying":
                    data["placenta_ll"] = lowlying_select(st, "nt_plac_ll")
            b1, b2 = st.columns(2)
            data["uterine_notching"] = b1.selectbox("Uterine artery notching", ["No notching", "Unilateral notching", "Bilateral notching"], key="early_notch")
            data["conception"] = b2.selectbox("Conception (UtA PI reference)", ["Natural", "IVF"], key="early_conception")
            st.caption("Uterine artery percentile is worked out from the CRL (valid for CRL 45-84 mm).")

    elif template_id == "NT_NB":
        with st.expander("Typist Inputs (NT / NB)", expanded=True):
            b1, b2, b3 = st.columns(3)
            data["cardiac_status"] = b1.selectbox("Cardiac pulsations", ["Present", "Absent"], key="nt_cardiac")
            st.session_state.setdefault("nt_cervix", "Normal")
            data["cervix"] = b2.text_input("Cervix length", key="nt_cervix")
            data["internal_os"] = b3.selectbox("Internal Os", ["Closed", "Open"], key="nt_os")
            b1, b2 = st.columns(2)
            data["subchorionic_present"] = b1.checkbox("Subchorionic collection present", key="nt_sch")
            if data["subchorionic_present"]:
                data["subchorionic_grade"] = b1.radio("Grade", ["Grade-I", "Grade-I/II", "Grade-II", "Grade-II/III"], horizontal=True, key="nt_sch_g")
            data["corpus_present"] = b2.checkbox("Corpus luteal cyst present", key="nt_cl")
            if data["corpus_present"]:
                d1, d2 = b2.columns(2)
                data["corpus_side"] = d1.selectbox("Side", ["Right", "Left"], key="nt_cl_s")
                data["corpus_size"] = d2.text_input("Size (mm)", key="nt_cl_z")
            b1, b2 = st.columns(2)
            _plac_opts = PLACENTA_OPTS
            data["placenta_pos"] = b1.selectbox("Placenta Position", _plac_opts, key="placenta_pos_widget")
            data["placenta_os"] = b2.selectbox("Placenta Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key="placenta_os_widget")
            if data["placenta_os"] == "Low-lying":
                data["placenta_ll"] = lowlying_select(st, "nt_plac_ll")
            data["choroid"] = st.selectbox("Choroid plexi / Falx", ["Appreciated, appear Normal", "Not well visualised"], key="nt_choroid")
            st.markdown("**First-trimester Doppler (percentiles are worked out from the CRL; valid for CRL 45-84 mm)**")
            b1, b2 = st.columns(2)
            data["uterine_notching"] = b1.selectbox("Uterine artery notching", ["No notching", "Unilateral notching", "Bilateral notching"], key="nt_notch")
            data["conception"] = b2.selectbox("Conception (UtA PI reference)", ["Natural", "IVF"], key="nt_conception")
            b1, b2 = st.columns(2)
            data["dv_pi"] = b1.text_input("Ductus Venosus PI", value="", key="nt_dvpi")
            data["dv_awave"] = b2.selectbox("DV A-Wave", ["Positive", "Absent", "Reversed"], key="nt_dvwave")

    elif template_id == IUD_OBS:
        with st.expander("IUD - Presentation & Findings", expanded=True):
            st.caption("Cardiac activity and fetal movements are always reported as absent in this template.")
            b1, b2 = st.columns(2)
            data["presentation"] = b1.selectbox("Presentation", PRESENTATION_OPTIONS,
                                                index=PRESENTATION_OPTIONS.index(PRESENTATION_DEFAULTS[IUD_OBS]),
                                                key=f"presentation_widget_{template_id}")
            data["lie"] = "None"
            if data["presentation"] in ("OBLIQUE LIE", "TRANSVERSE LIE"):
                data["head_side"] = b2.selectbox("Head in - Side", ["Right", "Left"], key="head_side_widget")
                data["head_level"] = b2.selectbox("Head in - Quadrant", ["Upper", "Lower"], key="head_level_widget")

            data["afi_method"] = st.radio("AFI Method", ["Largest Pocket", "Total 4 Quadrants"], horizontal=True, key="iud_afi_method")
            if data["afi_method"] == "Total 4 Quadrants":
                data["afi_total"] = st.text_input("AFI (cm)", key="iud_afi_total")
            else:
                data["afi_pocket"] = st.text_input("Largest Pocket (cm)", key="iud_afi_pocket")

            b1, b2 = st.columns(2)
            data["placenta_pos"] = b1.selectbox("Placenta Position", PLACENTA_OPTS, key="iud_plac_pos")
            data["placenta_os"] = b2.selectbox("Placenta Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key="iud_plac_os")
            if data["placenta_os"] == "Low-lying":
                data["placenta_ll"] = lowlying_select(st, "iud_plac_ll")
            elif data["placenta_os"] == "Placenta Praevia":
                data["praevia_grade"] = st.radio("Praevia Grade", ["Grade I", "Grade II", "Grade III", "Grade IV"], horizontal=True, key="iud_praevia")
            b1, b2 = st.columns(2)
            data["cervix"] = b1.text_input("Cervix Length (cm)", key="iud_cervix")
            data["internal_os"] = b2.selectbox("Internal Os", ["Closed", "Open"], key="iud_int_os")

            st.markdown("**Findings suggesting demise** (each appears in the body and in the impression)")
            f1, f2 = st.columns(2)
            for _i, (_k, _lbl, _b, _imp) in enumerate(IUD_FINDINGS):
                data[f"iud_{_k}"] = (f1 if _i % 2 == 0 else f2).checkbox(_lbl, key=f"sr_iud_{_k}_{v}")
            data["iud_other"] = st.text_input("Other finding (free text, e.g. Oligohydramnios noted)", key=f"sr_iud_other_{v}")

    elif not is_twin_template(template_id):
        with st.expander("Anatomy & Presentation", expanded=True):
            b1, b2 = st.columns(2)
            if not st.session_state.get("pres_user_set"):          # not chosen by the doctor yet: follow the template default
                st.session_state["presentation_widget"] = PRESENTATION_DEFAULTS.get(template_id, "CEPHALIC")
            elif "presentation_widget" not in st.session_state and st.session_state.get("pres_saved") in PRESENTATION_OPTIONS:
                st.session_state["presentation_widget"] = st.session_state["pres_saved"]   # Streamlit forgets widgets not shown (other template): restore
            data["presentation"] = b1.selectbox("Presentation", PRESENTATION_OPTIONS,
                                                key="presentation_widget", on_change=_mark_pres_set)
            st.session_state["pres_saved"] = data["presentation"]
            if data["presentation"] == "CEPHALIC":
                data["lie"] = b2.selectbox("Lie", LIE_OPTIONS, format_func=lambda c: LIE_LABELS.get(c, c), key="lie_widget")
            else:
                b2.selectbox("Lie", ["None"], disabled=True, key="lie_widget_disabled")  # lie options only apply to cephalic
                data["lie"] = "None"
            if data["presentation"] in ("OBLIQUE LIE", "TRANSVERSE LIE"):
                q1, q2 = st.columns(2)
                data["head_side"] = q1.selectbox("Head in - Side", ["Right", "Left"], key="head_side_widget")
                data["head_level"] = q2.selectbox("Head in - Quadrant", ["Upper", "Lower"], key="head_level_widget")

            data["afi_method"] = st.radio("AFI Method", ["Largest Pocket", "Total 4 Quadrants"], horizontal=True, key="sr_ocr_afi_method")
            if data["afi_method"] == "Total 4 Quadrants":
                data["afi_total"] = st.text_input("AFI (cm)", key="sr_ocr_afi_total")
            else:
                data["afi_pocket"] = st.text_input("Largest Pocket (cm)", key="sr_ocr_afi_pocket")

            b1, b2 = st.columns(2)
            data["placenta_pos"] = b1.selectbox("Placenta Position", PLACENTA_OPTS, key="sr_ocr_plac_pos")
            data["placenta_os"] = b2.selectbox("Placenta Relation to Os", ["away from the Os", "Low-lying", "Placenta Praevia"], key="sr_ocr_plac_os")

            if data["placenta_os"] == "Low-lying":
                data["placenta_ll"] = lowlying_select(st, "placenta_ll_widget")
                if data["placenta_ll"] == PLAC_LL_OPTS[0]:
                    data["placenta_dist"] = st.text_input("Distance from Internal Os (cm) [Optional]")
            elif data["placenta_os"] == "Placenta Praevia":
                data["praevia_grade"] = st.radio("Praevia Grade", ["Grade I", "Grade II", "Grade III", "Grade IV"], horizontal=True)

            b1, b2 = st.columns(2)
            data["cervix"] = b1.text_input("Cervix Length (cm)", key="sr_ocr_cervix")
            data["internal_os"] = b2.selectbox("Internal Os", ["Closed", "Open"])

            if template_id != ANOMALY:        # cord loop is not part of the Anomaly Scan report
                st.markdown("**Cord Loop**")
                data["cord_around"] = st.checkbox("Single loop of cord around the neck")
                data["cord_straddling"] = st.checkbox("Single loop of cord straddling the neck")

            if template_id in ("T3", COLOR_DOPPLER):
                st.markdown("**Scar Thickness**")
                data["measure_scar"] = st.checkbox("Measure Scar Thickness", key="sr_ocr_measure_scar")
                if data["measure_scar"]:
                    data["scar_size"] = st.text_input("Scar Size (mm)", key="sr_ocr_scar_size")

            # --- Soft markers (keys carry the form version, so they clear with a new patient) ---
            st.markdown("**Soft Markers**")
            data["cpc"] = st.checkbox("Choroid plexus cyst", key=f"sr_cpc_{v}")
            if data["cpc"]:
                k1, k2 = st.columns(2)
                data["cpc_right"] = k1.checkbox("Right", key=f"sr_cpc_right_{v}")
                if data["cpc_right"]:
                    data["cpc_right_size"] = k1.text_input("Right cyst size (mm)", key=f"sr_cpc_right_size_{v}")
                data["cpc_left"] = k2.checkbox("Left", key=f"sr_cpc_left_{v}")
                if data["cpc_left"]:
                    data["cpc_left_size"] = k2.text_input("Left cyst size (mm)", key=f"sr_cpc_left_size_{v}")
                if not (data["cpc_right"] or data["cpc_left"]):
                    st.caption("Tick Right and/or Left.")
            data["eif"] = st.checkbox("Echogenic intracardiac focus (EIF)", key=f"sr_eif_{v}")
            if data["eif"]:
                e1, e2 = st.columns(2)
                data["eif_left"] = e1.checkbox("Left ventricle", key=f"sr_eif_left_{v}")
                data["eif_right"] = e2.checkbox("Right ventricle", key=f"sr_eif_right_{v}")
                if not (data["eif_left"] or data["eif_right"]):
                    st.caption("Tick Left and/or Right ventricle.")

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
    _is_doppler_tpl = template_id in (COLOR_DOPPLER, TWINS_DOPPLER)      # both Doppler templates: always pink, no rotation used
    _is_anomaly_tpl = template_id in (ANOMALY, TWINS_ANOMALY)
    _is_iud_tpl = template_id in (IUD_OBS, TWINS_IUD)
    _is_pink_fixed = _is_doppler_tpl or template_id == TWINS_NT_NB       # twin NT/NB is always pink too
    _color_idx = PINK_THEME_INDEX if _is_pink_fixed else theme_for_template(template_id, load_color_index())   # Color Doppler is always pink; NT/NB skips Blue
    data["color_theme"] = _color_idx
    docx_bytes = build_docx_bytes(data, template_id)
    
    # Generate a unique filename: PatientName_GA_DateTime
    import datetime as dt
    _now = dt.datetime.now().strftime("%d%b%Y_%H%M")
    
    # Clean up the name and GA to avoid invalid file characters
    _clean_name = str(data.get('patient_name', 'Unknown')).replace(" ", "_").replace("/", "_")
    _clean_ga = str(data.get('mean_age', 'UnknownGA')).replace(" ", "").replace("±", "")
    
    _filename = f"{'DOPPLER' if _is_doppler_tpl else 'ANOMALY' if _is_anomaly_tpl else 'IUD' if _is_iud_tpl else 'OB'}_{_clean_name}_{_clean_ga}_{_now}.docx"

    st.download_button(
        label="Download Word Document",
        data=docx_bytes,
        file_name=_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        on_click=None if (_is_pink_fixed or template_id == TWINS_ANOMALY) else (lambda _u=_color_idx: advance_color_index(_u))   # Color Doppler does not use up a rotation colour
    )
    if template_id == TWINS_ANOMALY:
        st.caption("Table colour: Blue (fixed for Twin Anomaly Scan reports; the colour rotation is not affected)")
    elif _is_pink_fixed:
        st.caption("Table colour: Pink (fixed for this report type; the colour rotation is not affected)")
    else:
        st.caption(f"Table colour for this report: {COLOR_THEMES[_color_idx][0]} "
                   f"(next report: {COLOR_THEMES[_next_enabled(_color_idx)][0]})")
