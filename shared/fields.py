"""
Field and block specification for OB reports.

Each field is a dict:
    id          unique identifier (used in templates and SR mapping)
    label       display label
    type        number | enum | text | date | auto
    unit        optional unit string
    source      sr | calc | type | auto
    templates   list of template IDs where this field appears
    optional    bool (default False)
    choices     for enum fields
    default     optional default value
    notes       free text
"""

TEMPLATES = {
    "EARLY_OBS":   "Early OBS scan",
    "T2_PRE20":    "Second trimester (≤22 weeks)",
    "T2_POST20":   "Second trimester (>22 weeks)",
    "T3":          "Third trimester",
    "LEVEL2":      "Anomaly / Level II scan",
    "OBS_DOPPLER": "OBS with colour Doppler",
}

FIELDS = [
    # --- Block 0: Patient header ---
    dict(id="patient_name", label="Name", type="text", source="type",
         templates=list(TEMPLATES.keys())),
    dict(id="patient_age",  label="Age", type="number", unit="yr", source="type",
         templates=list(TEMPLATES.keys())),
    dict(id="report_date",  label="Date", type="date", source="auto",
         templates=list(TEMPLATES.keys())),
    dict(id="referred_by",  label="Ref. by", type="text", source="type",
         templates=list(TEMPLATES.keys())),

    # --- Block 1: Title and LMP ---
    dict(id="lmp", label="LMP", type="date", source="type",
         templates=list(TEMPLATES.keys()),
         choices=["date", "NOT SURE", "NOT KNOWN"]),

    # --- Block 2: Fetal status intro ---
    dict(id="presentation", label="Presentation", type="enum", source="type",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER"],
         choices=["variable", "cephalic", "breech", "transverse"],
         default="variable"),
    dict(id="lie", label="Lie", type="enum", source="type",
         templates=["T3", "OBS_DOPPLER"],
         choices=["LOA", "ROA", "LOP", "ROP"], optional=True),

    # --- Block 3: Early OBS table ---
    dict(id="gsac_status", label="G-Sac", type="enum", source="type",
         templates=["EARLY_OBS"],
         choices=["seen", "empty", "not seen"]),
    dict(id="msd", label="MSD", type="number", unit="mm", source="sr",
         templates=["EARLY_OBS"], optional=True),
    dict(id="crl", label="CRL", type="number", unit="mm", source="sr",
         templates=["EARLY_OBS"], optional=True),
    dict(id="yolk_sac", label="Yolk sac", type="text", source="type",
         templates=["EARLY_OBS"], optional=True),
    dict(id="cardiac_activity", label="Cardiac pulsations", type="enum",
         source="type", templates=["EARLY_OBS"],
         choices=["present", "not yet appreciable",
                  "not appreciated (multiple attempts)", "N/A"]),
    dict(id="fhr", label="FHR", type="number", unit="bpm", source="sr",
         templates=["EARLY_OBS", "T2_PRE20", "T2_POST20", "T3",
                    "LEVEL2", "OBS_DOPPLER"], optional=True),
    dict(id="subchorionic", label="Subchorionic collection", type="enum",
         source="type", templates=["EARLY_OBS"],
         choices=["absent", "grade I", "grade II", "grade III"]),
    dict(id="corpus_luteum", label="Corpus luteum cyst", type="text",
         source="type", templates=["EARLY_OBS"], optional=True),

    # --- Block 4: Biometry ---
    dict(id="bpd", label="BPD", type="number", unit="mm", source="sr",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"]),
    dict(id="hc",  label="HC",  type="number", unit="mm", source="sr",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"]),
    dict(id="ac",  label="AC",  type="number", unit="mm", source="sr",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"]),
    dict(id="fl",  label="FL",  type="number", unit="mm", source="sr",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"]),
    dict(id="hl",  label="HL",  type="number", unit="mm", source="sr",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER"]),

    # --- Block 5: Derived values ---
    dict(id="mean_age",     label="Mean age", type="text", source="calc",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER"]),
    dict(id="fetal_weight", label="Fetal weight", type="number", unit="g",
         source="sr", templates=["T2_PRE20", "T2_POST20", "T3",
                                 "OBS_DOPPLER", "LEVEL2"]),
    dict(id="afi",          label="AFI", type="text", source="type",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER"]),
    dict(id="edd_lmp",      label="EDD (LMP)", type="date", source="calc",
         templates=list(TEMPLATES.keys()), optional=True),
    dict(id="edd_cua",      label="EDD (CUA)", type="date", source="calc",
         templates=list(TEMPLATES.keys()), optional=True),

    # --- Block 6: Doppler ---
    dict(id="mean_uterine_pi", label="Mean uterine artery PI", type="number",
         source="sr", templates=["T2_PRE20"]),
    dict(id="right_uterine_pi", label="Right uterine artery PI", type="number",
         source="sr", templates=["T2_POST20", "T3"]),
    dict(id="umbilical_pi", label="Umbilical artery PI", type="number",
         source="sr", templates=["T2_POST20", "T3"]),

    # Full Doppler table (OBS_DOPPLER + LEVEL2)
    dict(id="dop_umb_sd",  label="Umbilical S/D", type="number", source="sr",
         templates=["OBS_DOPPLER"]),
    dict(id="dop_umb_ri",  label="Umbilical RI",  type="number", source="sr",
         templates=["OBS_DOPPLER"]),
    dict(id="dop_umb_pi",  label="Umbilical PI",  type="number", source="sr",
         templates=["OBS_DOPPLER"]),
    dict(id="dop_mca_pi",  label="MCA PI",        type="number", source="sr",
         templates=["OBS_DOPPLER"]),
    dict(id="dop_mca_psv", label="MCA PSV",       type="number", unit="cm/s",
         source="sr", templates=["OBS_DOPPLER"]),
    dict(id="dop_uta_pi",  label="Uterine (mean) PI", type="number",
         source="sr", templates=["OBS_DOPPLER"]),
    dict(id="dop_cpr",     label="CPR",           type="number", source="calc",
         templates=["OBS_DOPPLER"]),
    dict(id="dop_dv_pi",   label="Ductus venosus PI", type="number",
         source="type", templates=["OBS_DOPPLER", "LEVEL2"], optional=True),

    # --- Block 7: Anatomy / environment ---
    dict(id="placenta_position", label="Placenta position", type="enum",
         source="type",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"],
         choices=["anterior", "posterior", "fundal", "low-lying"]),
    dict(id="cervix_length", label="Cervix length", type="number", unit="cm",
         source="sr",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"]),
    dict(id="internal_os", label="Internal Os", type="enum", source="type",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER", "LEVEL2"],
         choices=["closed", "open"]),
    dict(id="cord_loop", label="Cord loop around neck", type="enum",
         source="type",
         templates=["T2_PRE20", "T2_POST20", "T3", "OBS_DOPPLER"],
         choices=["none seen", "present"], default="none seen"),
    dict(id="scar_thickness", label="Scar thickness", type="number",
         unit="mm", source="sr", templates=["T3"], optional=True),
]
