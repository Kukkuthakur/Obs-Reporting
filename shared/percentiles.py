import bisect

# ============================================================
# BARCELONA PERCENTILE CALCULATOR
# ============================================================

# --- Uterine Artery PI Reference (Gómez et al., FMB) ---
# Data extracted from search results [citation:2]
# Format: {GA_weeks: (5th_centile, 50th_centile, 95th_centile)}
UTA_PI_REF = {
    11: (1.18, 1.79, 2.70), 12: (1.11, 1.68, 2.53), 13: (1.05, 1.58, 2.38),
    14: (0.99, 1.49, 2.24), 15: (0.94, 1.41, 2.11), 16: (0.89, 1.33, 1.99),
    17: (0.85, 1.27, 1.88), 18: (0.81, 1.20, 1.79), 19: (0.78, 1.15, 1.70),
    20: (0.74, 1.10, 1.61), 21: (0.71, 1.05, 1.54), 22: (0.69, 1.00, 1.47),
    23: (0.66, 0.96, 1.41), 24: (0.64, 0.93, 1.35), 25: (0.62, 0.89, 1.30),
    26: (0.60, 0.86, 1.25), 27: (0.58, 0.84, 1.21), 28: (0.56, 0.81, 1.17),
    29: (0.55, 0.79, 1.13), 30: (0.54, 0.77, 1.10), 31: (0.52, 0.75, 1.06),
    32: (0.51, 0.73, 1.04), 33: (0.50, 0.71, 1.01), 34: (0.50, 0.70, 0.99),
    35: (0.49, 0.69, 0.97), 36: (0.48, 0.68, 0.95), 37: (0.48, 0.67, 0.94),
    38: (0.47, 0.66, 0.92), 39: (0.47, 0.65, 0.91), 40: (0.47, 0.65, 0.90),
    41: (0.47, 0.65, 0.89)
}

# --- Umbilical Artery PI Reference (from NHS table) ---
# Data from search results [citation:5]
# Format: {GA_weeks: (50th_centile, 95th_centile)}
UA_PI_REF = {
    19: (1.25, 1.63), 20: (1.22, 1.59), 21: (1.19, 1.55), 22: (1.17, 1.52),
    23: (1.14, 1.48), 24: (1.12, 1.47), 25: (1.09, 1.44), 26: (1.06, 1.41),
    27: (1.03, 1.38), 28: (1.00, 1.35), 29: (0.98, 1.32), 30: (0.95, 1.29),
    31: (0.93, 1.27), 32: (0.90, 1.25)
}

def _interpolate_percentile(value, ref_tuple):
    """Estimates percentile by linear interpolation between 5th, 50th, 95th."""
    p5, p50, p95 = ref_tuple
    if value <= p5:
        return 5.0
    if value >= p95:
        return 95.0
    if value <= p50:
        # Interpolate between 5 and 50
        return 5.0 + 45.0 * (value - p5) / (p50 - p5)
    else:
        # Interpolate between 50 and 95
        return 50.0 + 45.0 * (value - p50) / (p95 - p50)

def get_uterine_pi_percentile(pi_value, ga_weeks):
    """Returns estimated percentile for Uterine Artery PI."""
    if ga_weeks not in UTA_PI_REF:
        return None
    ref = UTA_PI_REF[ga_weeks]
    return round(_interpolate_percentile(pi_value, ref), 1)

def get_umbilical_pi_percentile(pi_value, ga_weeks):
    """Returns estimated percentile for Umbilical Artery PI."""
    if ga_weeks not in UA_PI_REF:
        return None
    p50, p95 = UA_PI_REF[ga_weeks]
    # Simple estimation based on 50th and 95th
    if pi_value >= p95:
        return 95.0
    if pi_value <= p50:
        return 50.0  # We only have 50th and 95th, so we can't be precise below 50
    return round(50.0 + 45.0 * (pi_value - p50) / (p95 - p50), 1)

def get_mca_pi_percentile(pi_value, ga_weeks):
    """Placeholder for Barcelona MCA PI percentiles."""
    # Add Barcelona MCA tables here
    return None

def get_cpr_percentile(cpr_value, ga_weeks):
    """Placeholder for Barcelona CPR percentiles."""
    # Add Barcelona CPR tables here
    return None
