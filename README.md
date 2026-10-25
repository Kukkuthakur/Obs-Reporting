# OB Reporting App

Obstetric ultrasound reporting application with DICOM SR auto-population.
Built with Streamlit (UI) + pynetdicom (DICOM listener).

## Status

Early development — validating report templates and field logic before
integrating live DICOM SR feed from the ultrasound machine.

## Architecture

Two independent entry points sharing common field definitions:

- `app.py` — Streamlit reporting interface (deployable to Streamlit Cloud)
- `dicom_listener.py` — DICOM Storage SCP (runs only on a local/LAN machine)
- `shared/` — field definitions, template specs, percentile calculators

**Important:** Streamlit Cloud cannot receive inbound DICOM connections
(no open TCP port, app sleeps). DICOM listening always runs on a local
or intranet machine. The deployed Streamlit app reads pre-parsed data
from local storage.

## Repository structure

    ob-reporting-app/
    ├── app.py                  # Streamlit UI entry point
    ├── dicom_listener.py       # DICOM Storage SCP entry point
    ├── requirements.txt
    ├── README.md
    ├── shared/
    │   ├── __init__.py
    │   ├── fields.py           # Field & block definitions (spec)
    │   ├── templates.py        # Report template composition rules
    │   └── percentiles.py      # Barcelona reference calculators
    ├── data/
    │   ├── received/           # DICOM SR files land here
    │   └── reports/            # Generated report data (JSON)
    └── config/
        └── dicom.yaml          # AE Title, port, storage paths

## Local development

    # Install dependencies
    pip install -r requirements.txt

    # Run the Streamlit UI (browser opens at http://localhost:8501)
    streamlit run app.py

    # In a separate terminal, run the DICOM listener
    python dicom_listener.py

    # Verify listener is up
    # From another machine: ping the host, then have the USG "Verify" the AE

## Streamlit Cloud deployment

1. Push repo to GitHub.
2. On share.streamlit.io, connect the repo, set main file to `app.py`.
3. Add any secrets (none needed yet) under app settings.
4. Deploy.

Note: app will sleep after 12h without traffic on the free tier.

## Configuration (config/dicom.yaml)

    ae_title: OB_REPORT_APP
    port: 11112
    storage_dir: data/received
    log_level: INFO

The AE Title and port here must match what is configured on the
ultrasound machine's DICOM destination settings.

## Field specification

See `shared/fields.py` — single source of truth for every field used
across all report templates. The template engine and the SR mapping
both read from this.

## Roadmap

- [x] Field & block specification (v0.2)
- [ ] Template engine (compose report from field values)
- [ ] Streamlit UI (template selector, field entry, preview)
- [ ] Barcelona percentile calculator (weight, umbilical, uterine, MCA, CPR)
- [ ] DICOM listener (receive SR, dump contents)
- [ ] SR parser (extract measurement nodes to field dict)
- [ ] Field mapping (SR concept name → field ID)
- [ ] Auto-populate UI from received SR
- [ ] PDF/DOCX export
- [ ] Twin support (deferred)
