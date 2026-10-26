"""
Field and block specification for OB reports.
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
    # Block 0: Patient header
    dict(id="patient_name", label="Name", type="text", source="type"),
    dict(id="patient_age", label="Age (yrs)", type="text", source="type"),
    dict(id="report_date", label="Date", type="text", source="type"),
    dict(id="referred_by", label="Referred By", type="text", source="type"),
    dict(id="lmp", label="LMP", type="text", source="type"),

    # Block 2: Intro
    dict(id="presentation", label="Presentation", type="enum", source="type",
         choices=["variable", "cephalic", "breech", "transverse"]),
    dict(id="lie", label="Lie", type="enum", source="type",
         choices=["LOA", "ROA", "LOP", "ROP"]),

    # Block 3: Early OBS
    dict(id="gsac_status", label="G-Sac", type="enum", source="type",
         choices=["seen", "empty", "not seen"]),
    dict(id="msd", label="MSD", type="text", source="sr"),
    dict(id="crl", label="CRL", type="text", source="sr"),
    dict(id="yolk_sac", label="Yolk sac", type="text", source="type"),
    dict(id="cardiac_activity", label="Cardiac pulsations", type="enum",
         choices=["present", "not yet appreciable", "N/A"]),
    dict(id="fhr", label="FHR (bpm)", type="text", source="sr"),

    # Block 4: Biometry
    dict(id="bpd", label="BPD (mm)", type="text", source="sr"),
    dict(id="hc", label="HC (mm)", type="text", source="sr"),
    dict(id="ac", label="AC (mm)", type="text", source="sr"),
    dict(id="fl", label="FL (mm)", type="text", source="sr"),
    dict(id="hl", label="HL (mm)", type="text", source="sr"),

    # Block 5: Derived values
    dict(id="mean_age", label="Mean Age", type="text", source="calc"),
    dict(id="fetal_weight", label="Fetal Weight (g)", type="text", source="sr"),
    dict(id="afi", label="AFI (Largest pocket cm)", type="text", source="type"),
    dict(id="edd_lmp", label="EDD (LMP)", type="text", source="calc"),
    dict(id="edd_cua", label="EDD (CUA)", type="text", source="calc"),

    # Block 6: Doppler
    dict(id="mean_uterine_pi", label="Mean Uterine PI", type="text", source="sr"),
    dict(id="right_uterine_pi", label="Right Uterine PI", type="text", source="sr"),
    dict(id="umbilical_pi", label="Umbilical PI", type="text", source="sr"),

    # Block 7: Anatomy
    dict(id="placenta_position", label="Placenta", type="enum", source="type",
         choices=["anterior", "posterior", "fundal"]),
    dict(id="cervix_length", label="Cervix Length (cm)", type="text", source="sr"),
    dict(id="internal_os", label="Internal Os", type="enum", source="type",
         choices=["closed", "open"]),
    dict(id="cord_loop", label="Cord loop", type="enum", source="type",
         choices=["none seen", "present"]),
    dict(id="scar_thickness", label="Scar Thickness (mm)", type="text", source="sr"),
]

FIELD_BY_ID = {f["id"]: f for f in FIELDS}
