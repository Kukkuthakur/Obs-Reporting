"""
Report template composition rules. Each template lists the blocks
that appear and in what order. Placeholders are field IDs from fields.py.
"""

TEMPLATES = {
    "EARLY_OBS":   "Early OBS scan",
    "T2_PRE20":    "Second trimester (≤22 weeks)",
    "T2_POST20":   "Second trimester (>22 weeks)",
    "T3":          "Third trimester",
    "LEVEL2":      "Anomaly / Level II scan",
    "OBS_DOPPLER": "OBS with colour Doppler",
}

BLOCKS = {
    "HEADER":   ["patient_name", "patient_age", "report_date", "referred_by"],
    "LMP":      ["lmp"],
    "INTRO":    ["presentation", "lie"],
    "EARLY":    ["gsac_status", "msd", "crl", "yolk_sac", "cardiac_activity",
                 "fhr", "subchorionic", "corpus_luteum"],
    "BIOMETRY": ["bpd", "hc", "ac", "fl", "hl"],
    "DERIVED":  ["mean_age", "fetal_weight", "afi", "edd_lmp", "edd_cua"],
    "DOPPLER_SIMPLE": ["mean_uterine_pi", "right_uterine_pi", "umbilical_pi"],
    "DOPPLER_FULL":   ["dop_umb_sd", "dop_umb_ri", "dop_umb_pi",
                       "dop_mca_pi", "dop_mca_psv", "dop_uta_pi",
                       "dop_cpr", "dop_dv_pi"],
    "ANATOMY":  ["placenta_position", "cervix_length", "internal_os",
                 "cord_loop"],
    "SCAR":     ["scar_thickness"],
}

TEMPLATE_BLOCKS = {
    "EARLY_OBS":   ["HEADER", "LMP", "EARLY", "DERIVED"],
    "T2_PRE20":    ["HEADER", "LMP", "INTRO", "BIOMETRY", "DERIVED",
                    "DOPPLER_SIMPLE", "ANATOMY"],
    "T2_POST20":   ["HEADER", "LMP", "INTRO", "BIOMETRY", "DERIVED",
                    "DOPPLER_SIMPLE", "ANATOMY"],
    "T3":          ["HEADER", "LMP", "INTRO", "BIOMETRY", "DERIVED",
                    "DOPPLER_SIMPLE", "ANATOMY", "SCAR"],
    "LEVEL2":      ["HEADER", "LMP", "INTRO", "BIOMETRY", "DERIVED",
                    "DOPPLER_SIMPLE", "DOPPLER_FULL", "ANATOMY"],
    "OBS_DOPPLER": ["HEADER", "LMP", "INTRO", "BIOMETRY", "DERIVED",
                    "DOPPLER_FULL", "ANATOMY"],
}
