"""
Radiology Report Generator — USG Whole Abdomen
v3.3.6-stable (Final Build)

Complete release: CLD, SAIO, Acalculus Cholecystitis, Reactive Mesenteric
Lymphadenitis templates; grouped Additional Findings; reworked Appendix;
Peri-Pancreatic/Peri-Portal nodes; structured Bowel Wall Thickening;
Uterus auto-append; Hindi app title.
"""

import io
import os
import re
import html
import hashlib
import sqlite3
from datetime import datetime
from zoneinfo import ZoneInfo

IST = ZoneInfo("Asia/Kolkata")


def _now_ist():
    """Return the current time in IST, tz-aware. Use this everywhere a
    timestamp will be shown to the user or written into a filename."""
    return datetime.now(IST)

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# CONFIG
# ============================================================

DB_PATH = os.path.join(os.path.expanduser("~"), "RadiologyReports", "reports.db")

PAGE_TOP_MARGIN = 3.2   # reduced by ~0.8cm so the demographics box
                        # sits two lines higher while still keeping
                        # 2 blank 11pt lines above it.
PAGE_BOTTOM_MARGIN = 3.0
PAGE_LEFT_MARGIN = 2.0
PAGE_RIGHT_MARGIN = 2.0

FONT_BODY = "Calibri"
FONT_SIZE_BODY = 11
FONT_SIZE_HALF = 5.5
FONT_SIZE_TITLE = 13
FONT_SIZE_IMPRESSION = 12
FONT_SIZE_DISCLAIMER = 7

TITLE_COLOR = RGBColor(0x1F, 0x4E, 0x79)

DISCLAIMER_TEXT = (
    "THIS IS ONLY A PROFESSIONAL OPINION NOT THE FINAL DIAGNOSIS. IT SHOULD BE "
    "CORRELATED CLINICALLY AND WITH OTHER RELEVANT INVESTIGATIONS. NOT VALID "
    "FOR MEDICO LEGAL PURPOSES."
)

IMPRESSION_NORMAL_FEMALE = [
    "NO FREE FLUID, LYMPHADENOPATHY, OBVIOUS BOWEL WALL THICKENING OR ANY SIGN OF INFLAMMATION SEEN AT THE TIME OF SCAN.",
    "UNREMARKABLE ABDOMEN & PELVIC SCAN.",
]
IMPRESSION_NORMAL_MALE = [
    "NO FREE FLUID, LYMPHADENOPATHY, OBVIOUS BOWEL WALL THICKENING OR ANY SIGN OF INFLAMMATION SEEN AT THE TIME OF SCAN.",
    "UNREMARKABLE ABDOMEN SCAN.",
]
IMPRESSION_REST_UNREMARKABLE = "REST OF THE ABDOMEN SCAN IS UNREMARKABLE."


# ============================================================
# CLINICAL TEMPLATES
# ============================================================

TEMPLATES = {
    "Chronic Liver Disease / Portal HTN": {
        "liver_size_num": 0,
        "liver_outline": "crenated",
        "liver_echo": "coarse",
        "liver_ihbr": "normal",
        "liver_portal": "normal",
        "gb_status": "adequately_distended",
        "gb_calculi": "none",
        "gb_sludge": "none",
        "gb_wall_check": True,
        "gb_wall_mm": "6.5",
        "cbd_status": "normal",
        "cbd_calc": False,
        "sp_size_num": 0,
        "sp_portal_mm": "",
        "sp_focal": "none",
        "kd_r_status": "normal",
        "kd_l_status": "normal",
        "ub_partially_empty": True,
        "ff_kind": "ascites",
        "ff_grade": "moderate",
        "bw_wall_thick": False,
        "impression_override": (
            "CHRONIC LIVER DISEASE WITH MODERATE ASCITES, MILD SPLENOMEGALY "
            "AND DILATED PORTAL VEIN SUGGESTING PORTAL HYPERTENSION."
        ),
        "impression_advice": "Adv- LFT, PT/INR & Clinical Correlation."
    },
    "Subacute Intestinal Obstruction": {
        "saio_subheading": "gaseous",
        "saio_distension_type": "liquid_filled",
        "saio_terminal_ileum": "25",
        "saio_ascending_colon": "39",
        "ff_kind": "inter_bowel",
        "ff_grade": "mild",
        "bw_wall_thick": False,
    },
    "Acalculus Cholecystitis / ?Acute Hepatitis": {
        "liver_size_num": 123,
        "liver_outline": "normal",
        "liver_echo": "raised",
        "liver_ihbr": "normal",
        "liver_portal": "normal",
        "gb_status": "adequately_distended",
        "gb_wall_check": True,
        "gb_wall_mm": "4.6",
        "gb_wall_edematous": True,
        "gb_wall_onion_peel": True,
        "gb_calculi": "none",
        "gb_sludge": "none",
        "cbd_status": "normal",
        "cbd_calc": False,
        "pr_cc": "45",
        "pr_median_lobe": True,
        "pr_median_lobe_size": "17",
        "pe_status": "present",
        "pe_laterality": "right",
        "pe_grade_right": "mild",
        "ff_kind": "ascites",
        "ff_grade": "mild",
        "ac_hepatitis_template": True,
    },   
}


# ============================================================
# SESSION-STATE MIRROR
# ============================================================

def wget(widget_key, default):
    if widget_key in st.session_state:
        v = st.session_state[widget_key]
        st.session_state[f"_mirror_{widget_key}"] = v
        return v
    return st.session_state.get(f"_mirror_{widget_key}", default)


def ss(widget_key, default):
    v = st.session_state.get(f"_mirror_{widget_key}")
    if v is not None:
        return v
    v = st.session_state.get(widget_key)
    if v is not None:
        return v
    return default


def _idx(widget_key, options, default):
    v = ss(widget_key, default)
    try:
        return options.index(v)
    except (ValueError, AttributeError):
        return 0


def _seed(widget_key, default):
    return ss(widget_key, default)


def _q(text):
    if not isinstance(text, str):
        return text
    return re.sub(r"\?\s+", "?", text)


# ============================================================
# HELPERS
# ============================================================

def capitalize_sentences(text):
    if not text:
        return text
    out = []
    capitalize_next = True
    for ch in text:
        if capitalize_next and ch.isalpha():
            out.append(ch.upper())
            capitalize_next = False
        else:
            out.append(ch)
            if ch in ".!?\n":
                capitalize_next = True
    return "".join(out)


def parse_kidney_length_mm(size_text):
    if not size_text:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", str(size_text))
    return float(m.group(1)) if m else None


def normalize_kidney_size_text(size_text):
    if not size_text:
        return ""
    s = str(size_text).strip()
    if not s:
        return ""
    s = re.sub(r"\s*[xX]\s*", "x", s)
    s = re.sub(r"\s*[mM]{2}\s*$", "", s)
    return f"{s}mm"


def _first_number(text):
    m = re.search(r"(\d+(?:\.\d+)?)", str(text or ""))
    if not m:
        return None
    try:
        return float(m.group(1))
    except (ValueError, TypeError):
        return None


def _lower_mm_in_text(t):
    if not isinstance(t, str):
        return t
    return re.sub(r"\bMM\b", "mm", t)


def join_findings(primary, findings):
    findings = [f for f in findings if f]
    if not findings:
        return primary.strip().rstrip(".") if primary else ""
    chunks = []
    remaining = list(findings)
    if primary:
        first_chunk = [remaining.pop(0)]
        if remaining:
            first_chunk.append(remaining.pop(0))
        if remaining:
            first_chunk.append(remaining.pop(0))
        if len(first_chunk) == 1:
            chunks.append(f"{primary} WITH {first_chunk[0]}")
        elif len(first_chunk) == 2:
            chunks.append(f"{primary} WITH {first_chunk[0]} & {first_chunk[1]}")
        else:
            chunks.append(f"{primary} WITH {first_chunk[0]} & {first_chunk[1]}")
            chunks.append(f"ALSO THE {first_chunk[2]} IS SEEN")
            while remaining:
                nxt = [remaining.pop(0)]
                if remaining:
                    nxt.append(remaining.pop(0))
                if len(nxt) == 1:
                    chunks.append(f"ALSO THE {nxt[0]} IS SEEN")
                else:
                    chunks.append(f"ALSO THE {nxt[0]} & {nxt[1]} ARE SEEN")
    else:
        first_chunk = [remaining.pop(0)]
        if remaining:
            first_chunk.append(remaining.pop(0))
        if remaining:
            first_chunk.append(remaining.pop(0))
        if len(first_chunk) == 1:
            chunks.append(first_chunk[0])
        elif len(first_chunk) == 2:
            chunks.append(f"{first_chunk[0]} & {first_chunk[1]}")
        else:
            chunks.append(f"{first_chunk[0]} & {first_chunk[1]}")
            chunks.append(f"ALSO THE {first_chunk[2]} IS SEEN")
        while remaining:
            nxt = [remaining.pop(0)]
            if remaining:
                nxt.append(remaining.pop(0))
            if len(nxt) == 1:
                chunks.append(f"ALSO THE {nxt[0]} IS SEEN")
            else:
                chunks.append(f"ALSO THE {nxt[0]} & {nxt[1]} ARE SEEN")
    result = ". ".join(chunks)
    if not result.endswith("."):
        result += "."
    return result


# ============================================================
# PEDIATRIC TABLES
# ============================================================

PEDIATRIC_SPLEEN_MAX_MM = {
    (0.0, 0.25, "F"): 55,  (0.0, 0.25, "M"): 68,
    (0.25, 0.5, "F"): 56,  (0.25, 0.5, "M"): 70,
    (0.5, 1.0, "F"): 75,   (0.5, 1.0, "M"): 74,
    (1.0, 2.0, "F"): 82,   (1.0, 2.0, "M"): 83,
    (2.0, 4.0, "F"): 89,   (2.0, 4.0, "M"): 99,
    (4.0, 6.0, "F"): 95,   (4.0, 6.0, "M"): 99,
    (6.0, 8.0, "F"): 100,  (6.0, 8.0, "M"): 105,
    (8.0, 10.0, "F"): 105, (8.0, 10.0, "M"): 112,
    (10.0, 12.0, "F"): 114,(10.0, 12.0, "M"): 113,
    (12.0, 14.0, "F"): 116,(12.0, 14.0, "M"): 117,
    (14.0, 18.0, "F"): 110,(14.0, 18.0, "M"): 125,
}

PEDIATRIC_LIVER_MAX_MM = {
    (0.0, 0.25): 90, (0.25, 0.5): 95, (0.5, 0.75): 100, (0.75, 1.0): 100,
    (1.0, 2.5): 105, (2.5, 3.0): 105, (3.0, 5.0): 115, (5.0, 7.0): 125,
    (7.0, 9.0): 130, (9.0, 11.0): 135, (11.0, 13.0): 140, (13.0, 15.0): 140,
    (15.0, 18.0): 145,
}


def _parse_mm_value(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", "."))
    except ValueError:
        return None


def parse_age(age_text):
    if age_text is None:
        return None
    txt = str(age_text).strip()
    if not txt:
        return None
    match = re.search(r"(\d+(\.\d+)?)", txt)
    if not match:
        return None
    try:
        return float(match.group(1))
    except (ValueError, TypeError):
        return None


def get_pediatric_liver_max_mm(age_years):
    for (lo, hi), max_mm in PEDIATRIC_LIVER_MAX_MM.items():
        if lo <= age_years < hi:
            return max_mm
    return None


def get_pediatric_spleen_max_mm(age_years, sex):
    for (lo, hi, s), max_mm in PEDIATRIC_SPLEEN_MAX_MM.items():
        if s == sex and lo <= age_years < hi:
            return max_mm
    return None


def classify_liver_adult(size_mm):
    if size_mm < 150:
        return "normal"
    elif size_mm < 151:
        return "borderline"
    elif size_mm <= 170:
        return "mild"
    elif size_mm <= 190:
        return "moderate"
    else:
        return "gross"


def classify_spleen_adult(size_mm):
    if size_mm < 120:
        return "normal"
    elif size_mm < 121:
        return "borderline"
    elif size_mm < 140:
        return "mild"
    elif size_mm < 141:
        return "mild_to_moderate"
    elif size_mm <= 160:
        return "moderate"
    elif size_mm <= 170:
        return "moderate_to_gross"
    else:
        return "gross"


def classify_portal_vein(pv_mm):
    v = _parse_mm_value(pv_mm)
    if v is None:
        return "unknown"
    if v <= 13.0:
        return "normal"
    elif v < 14.0:
        return "prominent"
    else:
        return "dilated"


def auto_classify_liver(size_mm, age_text, sex):
    try:
        size = float(size_mm)
    except (ValueError, TypeError):
        return "normal"
    if size <= 0:
        return "normal"
    age = parse_age(age_text)
    if age is None:
        return classify_liver_adult(size)
    if age < 18:
        max_mm = get_pediatric_liver_max_mm(age)
        if max_mm is None:
            return classify_liver_adult(size)
        return "enlarged_for_age" if size > max_mm else "normal"
    return classify_liver_adult(size)


def auto_classify_spleen(size_mm, age_text, sex):
    try:
        size = float(size_mm)
    except (ValueError, TypeError):
        return "normal"
    if size <= 0:
        return "normal"
    age = parse_age(age_text)
    if age is None:
        return classify_spleen_adult(size)
    if age < 18:
        max_mm = get_pediatric_spleen_max_mm(age, sex)
        if max_mm is None:
            return classify_spleen_adult(size)
        return "enlarged_for_age" if size > max_mm else "normal"
    return classify_spleen_adult(size)


LIVER_STATUS_LABELS = {
    "normal": "Normal size", "borderline": "Borderline enlarged",
    "mild": "Mildly enlarged", "moderate": "Moderately enlarged",
    "gross": "Grossly enlarged", "enlarged_for_age": "Enlarged for age",
}
SPLEEN_STATUS_LABELS = {
    "normal": "Normal size",
    "borderline": "Borderline enlarged",
    "mild": "Mildly enlarged",
    "mild_to_moderate": "Mild to moderately enlarged",
    "moderate": "Moderately enlarged",
    "moderate_to_gross": "Moderately to grossly enlarged",
    "gross": "Grossly enlarged",
    "enlarged_for_age": "Enlarged for age",
}
SPLEEN_IMPRESSION_LABELS = {
    "borderline": "BORDERLINE SPLENOMEGALY",
    "mild": "MILD SPLENOMEGALY",
    "mild_to_moderate": "MILD TO MODERATE SPLENOMEGALY",
    "moderate": "MODERATE SPLENOMEGALY",
    "moderate_to_gross": "MODERATE TO GROSS SPLENOMEGALY",
    "gross": "GROSS SPLENOMEGALY",
}

HEPATOSPLENOMEGALY_COMBINE = {
    "borderline": "BORDERLINE",
    "mild": "MILD",
    "moderate": "MODERATE",
    "gross": "GROSS",
    "mild_to_moderate": "MILD TO MODERATE",
    "moderate_to_gross": "MODERATE TO GROSS",
}


# ============================================================
# KIDNEY CONSTANTS
# ============================================================

URETER_LEVELS = {
    "renal_pelvis": ("in", "HYDRONEPHROSIS", "RENAL PELVIS",
                     "renal pelvis", "renal pelvis cal"),
    "puj": ("at", "HYDRONEPHROSIS", "PELVI-URETERIC JUNCTION",
            "pelvi-ureteric junction", "pelvi-ureteric junction cal"),
    "proximal_ureter": ("in", "HYDROURETERONEPHROSIS", "PROXIMAL URETERIC",
                        "proximal ureter", "proximal ureteric cal"),
    "mid_ureter": ("in", "HYDROURETERONEPHROSIS", "MID-URETERIC",
                   "mid-ureter", "mid-ureteric cal"),
    "distal_ureter": ("in", "HYDROURETERONEPHROSIS", "DISTAL URETERIC",
                      "distal ureter", "distal ureteric cal"),
    "vuj": ("at", "HYDROURETERONEPHROSIS", "VESICO-URETERIC JUNCTION",
            "vesico-ureteric junction", "vesico-ureteric junction cal"),
}

URETER_GRADES = {
    "none": None, "no_significant": "NO SIGNIFICANT",
    "minimal": "MINIMAL", "mild": "MILD", "moderate": "MODERATE",
}

URETER_GRADES_SENTENCE = {
    "no_significant": "no significant", "minimal": "minimal",
    "mild": "mild", "moderate": "moderate",
}

POLE_LABELS = {
    "upper": "upper-pole", "upper_mid": "upper-mid pole",
    "mid": "mid pole", "lower_mid": "lower-mid pole", "lower": "lower-pole",
}

POLE_OPTIONS = ["upper", "upper_mid", "mid", "lower_mid", "lower"]

SMALL_KIDNEY_MM = 82
SMALL_KIDNEY_MIN_AGE = 15

CALCULUS_TINY_MAX_MM = 3.0
CALCULUS_SMALL_MAX_MM = 4.0
CALCULUS_LARGE_MIN_MM = 10.0

CBD_DILATED_MM = 6.0


def calculus_size_qualifier(size_mm):
    """Return 'tiny' (≤3mm), 'small' (3.1–4mm), 'large' (>10mm), or ''."""
    v = _parse_mm_value(size_mm)
    if v is None:
        return ""
    if v <= CALCULUS_TINY_MAX_MM:
        return "tiny"
    if v <= CALCULUS_SMALL_MAX_MM:
        return "small"
    if v > CALCULUS_LARGE_MIN_MM:
        return "large"
    return ""
    

ECHO_GRADE_WORD = {
    "mildly_raised": "mildly raised",
    "moderately_raised": "moderately raised",
    "significantly_raised": "significantly raised",
}
ECHO_GRADE_WORD_UP = {
    "mildly_raised": "MILDLY RAISED",
    "moderately_raised": "MODERATELY RAISED",
    "significantly_raised": "SIGNIFICANTLY RAISED",
}
CMD_WORD = {"preserved": "preserved", "hazy": "hazy", "lost": "lost"}
CMD_WORD_UP = {"preserved": "PRESERVED", "hazy": "HAZY", "lost": "LOST"}


# ============================================================
# PROSTATE CONSTANTS
# ============================================================

PROSTATE_GRADE_BANDS = [
    (30, 36, "borderline bulky", "BORDERLINE PROSTATOMEGALY"),
    (36, 46, "mildly bulky", "GRADE-I PROSTATOMEGALY"),
    (46, 51, "mild to moderately bulky", "GRADE-I/II PROSTATOMEGALY"),
    (51, 71, "moderately bulky", "GRADE-II PROSTATOMEGALY"),
    (71, float("inf"), "moderate to grossly bulky", "GRADE-III/IV PROSTATOMEGALY"),
]


def prostate_band_from_volume(cc):
    try:
        v = float(cc)
    except (ValueError, TypeError):
        return None
    if v < 30:
        return None
    for lo, hi, body_desc, imp_label in PROSTATE_GRADE_BANDS:
        if lo <= v < hi:
            return body_desc, imp_label
    return None


# ============================================================
# APPENDIX CONSTANTS
# ============================================================

APPENDIX_NORMAL_MAX_MM = 6.0
APPENDIX_PROMINENT_MAX_MM = 7.0


def classify_appendix(diameter_mm):
    v = _parse_mm_value(diameter_mm)
    if v is None:
        return None
    if v <= APPENDIX_NORMAL_MAX_MM:
        return "normal"
    if v <= APPENDIX_PROMINENT_MAX_MM:
        return "prominent"
    return "dilated"


# ============================================================
# UTERUS CONSTANTS
# ============================================================

FIBROID_TYPE_LABELS = {
    "intramural": ("an intramural fibroid", "FIGO-4"),
    "intramural_subserosal": ("an intramural>subserosal fibroid", "FIGO-4/5"),
    "intramural_submucosal": ("an intramural>submucosal fibroid", "FIGO-4/3"),
    "submucosal_intramural": ("a submucosal>intramural fibroid", "FIGO-2/3"),
    "subserosal_intramural":
        ("a partially extrusive subserosal>intramural fibroid", "FIGO-5/4"),
}
FIBROID_TYPE_LABELS_UP = {
    "intramural": ("AN INTRAMURAL FIBROID", "FIGO-4"),
    "intramural_subserosal": ("AN INTRAMURAL>SUBSEROSAL FIBROID", "FIGO-4/5"),
    "intramural_submucosal": ("AN INTRAMURAL>SUBMUCOSAL FIBROID", "FIGO-4/3"),
    "submucosal_intramural": ("A SUBMUCOSAL>INTRAMURAL FIBROID", "FIGO-2/3"),
    "subserosal_intramural":
        ("A PARTIALLY EXTRUSIVE SUBSEROSAL>INTRAMURAL FIBROID", "FIGO-5/4"),
}
PARTIALLY_EXTRUSIVE_TYPES = {"intramural_subserosal", "subserosal_intramural"}
FIBROID_LOC_OPTIONS = ["fundal", "anterior", "posterior", "lateral", "cervical"]
FIBROID_LATERALITY_OPTIONS = ["right", "left"]
FIBROID_TYPE_OPTIONS = [
    ("intramural", "Intramural"),
    ("intramural_subserosal", "Intramural > subserosal"),
    ("intramural_submucosal", "Intramural > submucosal"),
    ("submucosal_intramural", "Submucosal > intramural"),
    ("subserosal_intramural", "Subserosal > intramural"),
]


def _fibroid_location_body(location, laterality=""):
    """Body phrase fragment for a fibroid lesion location."""
    if location == "lateral":
        side = laterality or "right"
        return f"the {side} lateral wall of uterine myometrium"
    return f"the {location} myometrium"


def _fibroid_location_impression(location, laterality=""):
    """Uppercase impression phrase fragment for a fibroid lesion location."""
    if location == "lateral":
        side = laterality or "right"
        return f"THE {side.upper()} LATERAL WALL"
    return f"THE {location.upper()} MYOMETRIUM"


def classify_uterus_size(length_mm):
    if length_mm is None:
        return "normal"
    if length_mm < 90:
        return "normal"
    if length_mm <= 100:
        return "bulky"
    if length_mm < 120:
        return "moderately_bulky"
    return "grossly_bulky"


UTERUS_SIZE_BODY_WORD = {
    "normal": None, "bulky": "bulky",
    "moderately_bulky": "moderately bulky",
    "grossly_bulky": "grossly bulky",
}
UTERUS_SIZE_IMPRESSION = {
    "normal": "", "bulky": "BULKY UTERUS",
    "moderately_bulky": "MODERATELY BULKY UTERUS",
    "grossly_bulky": "GROSSLY BULKY UTERUS",
}
UTERUS_SIZE_IMPRESSION_ADJ = {
    "normal": "", "bulky": "BULKY",
    "moderately_bulky": "MODERATELY BULKY",
    "grossly_bulky": "GROSSLY BULKY",
}

ENDOMETRIAL_COLLECTION_LABELS = {
    "none": "No collection seen in the endometrial cavity.",
    "mild": "Mild collection seen in the endometrial cavity.",
    "moderate": "Moderate collection seen in the endometrial cavity.",
}
ENDOMETRIAL_COLLECTION_HET = {
    "mild": "Mild heterogenous collection seen in the endometrial cavity.",
    "moderate": "Moderate heterogenous collection seen in the endometrial cavity.",
}
ENDOMETRIAL_COLLECTION_IMPRESSION = {
    "none": None, "mild": "MILD COLLECTION IN THE ENDOMETRIAL CANAL.",
    "moderate": "MODERATE COLLECTION IN THE ENDOMETRIAL CANAL.",
}
ENDOMETRIAL_COLLECTION_HET_IMPRESSION = {
    "mild": "MILD HETEROGENOUS COLLECTION IN THE ENDOMETRIAL CANAL.",
    "moderate": "MODERATE HETEROGENOUS COLLECTION IN THE ENDOMETRIAL CANAL.",
}


def classify_endometrium(mm):
    if mm is None:
        return None
    if mm < 2:
        return "thinned_out"
    if mm <= 15:
        return None
    if mm < 20:
        return "thickened"
    return "significantly_thickened"


ADENOMYOSIS_FEATURES = [
    "globular_shape", "asymmetric_posterior",
    "venetian_blind_sign", "subendometrial_cysts",
]
ADENOMYOSIS_INTERFACE_OPTIONS = [
    "none", "indistinct", "barely_perceptible_fundus", "lost",
]


# ============================================================
# OVARY CONSTANTS
# ============================================================

OVARY_LARGE_MM = 50
OVARY_BULKY_CC = 10.0
OVARY_FOLLICLE_MAX_MM = 30

OVARY_CYST_TYPES = {
    "simple_cyst": "simple cyst",
    "hemorrhagic_cyst": "hemorrhagic cyst",
    "hemorrhagic_follicle": "hemorrhagic follicle",
    "follicular_cyst": "follicular cyst",
    "corpus_luteum": "corpus luteum cyst",
}
OVARY_CYST_TYPES_UP = {
    "simple_cyst": "SIMPLE CYST",
    "hemorrhagic_cyst": "HEMORRHAGIC CYST",
    "hemorrhagic_follicle": "HEMORRHAGIC FOLLICLE",
    "follicular_cyst": "FOLLICULAR CYST",
    "corpus_luteum": "CORPUS LUTEUM CYST",
}
OVARY_CYST_TYPES_UP_PLURAL = {
    "simple_cyst": "SIMPLE CYSTS",
    "hemorrhagic_cyst": "HEMORRHAGIC CYSTS",
    "hemorrhagic_follicle": "HEMORRHAGIC FOLLICLES",
    "follicular_cyst": "FOLLICULAR CYSTS",
    "corpus_luteum": "CORPUS LUTEUM CYSTS",
}

PCOS_OUTCOME_CONCLUSIVE = "conclusive"
PCOS_OUTCOME_PARTIAL = "partial"
PCOS_OUTCOME_NOT_CONCLUSIVE = "not_conclusive"

PCOS_OUTCOME_PHRASES = {
    PCOS_OUTCOME_CONCLUSIVE: (
        "NEEDS CLINICO-LAB CORRELATION TO RULE OUT POLYCYSTIC OVARIAN "
        "SPECTRUM CHANGES (PCOS). Adv - LH/FSH & AMH Correlation."
    ),
    PCOS_OUTCOME_PARTIAL: (
        "PARTIALLY FITS ON THE POLYCYSTIC OVARIAN SPECTRUM. "
        "Adv- LH/FSH & AMH Correlation."
    ),
    PCOS_OUTCOME_NOT_CONCLUSIVE: (
        "FEATURES DO NOT APPEAR CONCLUSIVE FOR POLYCYSTIC OVARIAN SPECTRUM "
        "CHANGES (PCOS). Adv - LH/FSH & AMH Correlation."
    ),
}


def _effective_ovary_type(finding):
    ftype = finding.get("type", "")
    size = finding.get("size_mm", "")
    n = _first_number(size)
    if n is None:
        return ftype
    if n < OVARY_FOLLICLE_MAX_MM:
        if ftype == "simple_cyst":
            return "follicular_cyst"
        if ftype == "hemorrhagic_cyst":
            return "hemorrhagic_follicle"
    return ftype


# ============================================================
# MLN CONSTANTS
# ============================================================

MLN_COUNT_WORD = {"few": "Few", "multiple": "Multiple", "innumerable": "Innumerable"}
MLN_COUNT_WORD_UP = {"few": "FEW", "multiple": "MULTIPLE", "innumerable": "INNUMERABLE"}
MLN_SIZE_WORD = {
    "small": "small (SAD<7mm)",
    "enlarged": "enlarged (SAD>7mm)",
    "mixed": "small & enlarged (SAD>7mm)",
}
MLN_SIZE_WORD_UP = {
    "small": "SMALL (SAD<7MM)",
    "enlarged": "ENLARGED (SAD>7MM)",
    "mixed": "SMALL & ENLARGED (SAD>7MM)",
}
MLN_LOCATION_BODY = {
    "bilateral": " in the bilateral pre & para-aortic peri-umbilical region",
    "left": ", predominantly in the left para-aortic region",
    "right": ", predominantly in the right para-aortic region",
}
MLN_LOCATION_IMP = {
    "bilateral": " IN THE BILATERAL PRE & PARA-AORTIC PERI-UMBILICAL REGION",
    "left": ", PREDOMINANTLY IN THE LEFT PARA-AORTIC REGION",
    "right": ", PREDOMINANTLY IN THE RIGHT PARA-AORTIC REGION",
}

MLN_INFECTIVE_TAIL = (
    " - LIKELY INFECTIVE PROCESS(TUBERCULAR). "
    "Adv- Mantoux/CBNAAT, Lab & Clinical Correlation."
)


# ============================================================
# PPLN CONSTANTS
# ============================================================

PPLN_LOCATION_BODY = {
    "peri_pancreatic": "peri-pancreatic",
    "peri_portal": "peri-portal",
    "both": "peri-pancreatic & peri-portal",
}
PPLN_LOCATION_IMP = {
    "peri_pancreatic": "PERI-PANCREATIC",
    "peri_portal": "PERI-PORTAL",
    "both": "PERI-PANCREATIC & PERI-PORTAL",
}


# ============================================================
# FLUID / EFFUSION CONSTANTS
# ============================================================

EFF_LATERALITY_OPTS = ["right", "left", "bilateral", "bilateral_asym"]
EFF_GRADE_OPTS =
