"""
DICOM Storage SCP for receiving SR from the ultrasound machine.
Runs only on a local/LAN machine — not deployable to Streamlit Cloud.

Reads config from config/dicom.yaml. Saves incoming DICOM objects
to the storage directory and prints a summary.
"""

import logging
from pathlib import Path

import yaml
from pydicom import dcmwrite
from pynetdicom import AE, evt, StoragePresentationContexts

CONFIG_PATH = Path("config/dicom.yaml")


def load_config() -> dict:
    with CONFIG_PATH.open() as f:
        return yaml.safe_load(f)


def make_handler(storage_dir: Path):
    def handle_store(event):
        ds = event.dataset
        ds.file_meta = event.file_meta
        out = storage_dir / f"{ds.SOPInstanceUID}.dcm"
        dcmwrite(str(out), ds, enforce_file_format=True)
        logging.info(
            "Received %s — SOP Class %s — saved %s",
            ds.get("PatientName", "?"),
            ds.get("SOPClassUID", "?"),
            out.name,
        )
        return 0x0000
    return handle_store


def main() -> None:
    cfg = load_config()
    logging.basicConfig(
        level=cfg.get("log_level", "INFO"),
        format="%(asctime)s %(levelname)s %(message)s",
    )

    storage_dir = Path(cfg["storage_dir"])
    storage_dir.mkdir(parents=True, exist_ok=True)

    ae = AE(ae_title=cfg["ae_title"])
    ae.supported_contexts = StoragePresentationContexts

    handlers = [(evt.EVT_C_STORE, make_handler(storage_dir))]

    logging.info(
        "Starting DICOM listener — AE Title %s on port %s",
        cfg["ae_title"], cfg["port"],
    )
    ae.start_server(("0.0.0.0", int(cfg["port"])), evt_handlers=handlers)


if __name__ == "__main__":
    main()
