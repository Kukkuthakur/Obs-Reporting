"""
Barcelona / FMF reference calculators. Placeholder functions —
equations and lookup tables to be filled in from the Barcelona
reference framework. Each returns a percentile (0-100) given the
measured value and gestational age in days.
"""

def umbilical_pi_percentile(pi: float, ga_days: int) -> float:
    raise NotImplementedError

def uterine_pi_percentile(pi: float, ga_days: int) -> float:
    raise NotImplementedError

def mca_pi_percentile(pi: float, ga_days: int) -> float:
    raise NotImplementedError

def cpr_percentile(cpr: float, ga_days: int) -> float:
    raise NotImplementedError

def efw_percentile(efw_grams: float, ga_days: int) -> float:
    raise NotImplementedError
