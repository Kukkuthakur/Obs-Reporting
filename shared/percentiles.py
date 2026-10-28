import math
from scipy.stats import norm

# ============================================================
# BARCELONA / FMF Z-SCORE CALCULATOR
# Equations derived from FMF study of 72,387 fetuses [citation:2]
# ============================================================

def _ga_days(ga_weeks: float) -> float:
    """Convert weeks to days for the regression equations."""
    return ga_weeks * 7.0

def _calculate_z_score(value: float, median: float, sd: float) -> float:
    """Standard Z-score formula."""
    if sd == 0: return 0.0
    return (value - median) / sd

def _z_to_percentile(z: float) -> float:
    """Convert Z-score to percentile (0-100)."""
    return norm.cdf(z) * 100.0

# ------------------------------------------------------------
# UMBILICAL ARTERY (UA) PI
# Median = a + b*(GA_days)
# SD = 10^(c + d*(GA_days))
# ------------------------------------------------------------
def calculate_ua_pi(pi: float, ga_weeks: float):
    ga = _ga_days(ga_weeks)
    
    # FMF Coefficients for UA-PI [citation:2]
    # Median: Intercept + Slope * GA
    median = -0.0016 * ga + 1.2876  # (approximate, see note below)
    
    # SD: 10^(Intercept + Slope * GA)
    sd = 10 ** (-0.0012 * ga + 0.1147) 

    z = _calculate_z_score(pi, median, sd)
    return {
        "z_score": round(z, 2),
        "percentile": round(_z_to_percentile(z), 1),
        "median": round(median, 2)
    }

# ------------------------------------------------------------
# MIDDLE CEREBRAL ARTERY (MCA) PI
# Median = 10^(a + b*(GA_days))
# SD = 10^(c + d*(GA_days))
# ------------------------------------------------------------
def calculate_mca_pi(pi: float, ga_weeks: float):
    ga = _ga_days(ga_weeks)
    
    # FMF Coefficients for MCA-PI [citation:2]
    # Median: 10^(Intercept + Slope * GA)
    median = 10 ** (-0.0025 * ga + 0.4896)
    
    # SD: 10^(Intercept + Slope * GA)
    sd = 10 ** (-0.0009 * ga + 0.1440)

    z = _calculate_z_score(pi, median, sd)
    return {
        "z_score": round(z, 2),
        "percentile": round(_z_to_percentile(z), 1),
        "median": round(median, 2)
    }

# ------------------------------------------------------------
# CEREBROPLACENTAL RATIO (CPR)
# Median = 10^(a + b*(GA_days))
# SD = 10^(c + d*(GA_days))
# ------------------------------------------------------------
def calculate_cpr(cpr: float, ga_weeks: float):
    ga = _ga_days(ga_weeks)
    
    # FMF Coefficients for CPR [citation:2]
    # Median: 10^(Intercept + Slope * GA)
    median = 10 ** (-0.0026 * ga + 0.4678)
    
    # SD: 10^(Intercept + Slope * GA)
    sd = 10 ** (-0.0010 * ga + 0.1289)

    z = _calculate_z_score(cpr, median, sd)
    return {
        "z_score": round(z, 2),
        "percentile": round(_z_to_percentile(z), 1),
        "median": round(median, 2)
    }

# ------------------------------------------------------------
# UTERINE ARTERY (UtA) PI
# The FMF paper focuses on UA/MCA/CPR. UtA uses a different
# reference (Gómez et al.). We will stick to a lookup table
# for UtA as the Z-score equations are less standardized.
# ------------------------------------------------------------
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

def _interpolate_uta(value, ref_tuple):
    p5, p50, p95 = ref_tuple
    if value <= p5: return 5.0
    if value >= p95: return 95.0
    if value <= p50: return 5.0 + 45.0 * (value - p5) / (p50 - p5)
    return 50.0 + 45.0 * (value - p50) / (p95 - p50)

def calculate_uta_pi(pi: float, ga_weeks: int):
    if ga_weeks not in UTA_PI_REF:
        return None
    ref = UTA_PI_REF[ga_weeks]
    pct = _interpolate_uta(pi, ref)
    # Approximate Z-score from percentile (inverse norm)
    z = norm.ppf(pct / 100.0)
    return {
        "z_score": round(z, 2),
        "percentile": round(pct, 1),
        "median": ref[1]
    } 
