"""
Streamlit UI entry point for the OB reporting app.
"""

import streamlit as st

from shared.templates import TEMPLATES, TEMPLATE_BLOCKS, BLOCKS
from shared.fields import FIELDS

FIELD_BY_ID = {f["id"]: f for f in FIELDS}


def render_field(field: dict) -> None:
    label = field["label"]
    unit = field.get("unit")
    if unit:
        label = f"{label} ({unit})"

    key = field["id"]
    ftype = field["type"]

    if ftype == "enum":
        choices = field.get("choices", [])
        default = field.get("default", choices[0] if choices else None)
        st.selectbox(label, choices, index=choices.index(default)
                     if default in choices else 0, key=key)
    elif ftype == "number":
        st.number_input(label, value=0.0, step=0.1, key=key)
    elif ftype == "date":
        st.text_input(label, key=key, placeholder="DD-MM-YYYY")
    else:
        st.text_input(label, key=key)


def main() -> None:
    st.set_page_config(page_title="OB Reporting", layout="wide")
    st.title("OB Reporting App")

    template_id = st.sidebar.selectbox(
        "Report template",
        options=list(TEMPLATES.keys()),
        format_func=lambda k: TEMPLATES[k],
    )

    st.sidebar.caption(f"Template: {TEMPLATES[template_id]}")
    st.sidebar.markdown("---")
    st.sidebar.info(
        "DICOM SR auto-populate not yet wired. "
        "All fields currently manual."
    )

    st.header(TEMPLATES[template_id])

    for block_name in TEMPLATE_BLOCKS[template_id]:
        block_fields = BLOCKS[block_name]
        with st.expander(block_name.title(), expanded=True):
            cols = st.columns(2)
            for i, field_id in enumerate(block_fields):
                field = FIELD_BY_ID.get(field_id)
                if not field:
                    continue
                if field.get("optional") and not st.checkbox(
                    f"Include {field['label']}", key=f"inc_{field_id}"
                ):
                    continue
                with cols[i % 2]:
                    render_field(field)

    st.markdown("---")
    if st.button("Generate report preview"):
        st.info("Template engine not yet wired. Fields captured above.")


if __name__ == "__main__":
    main()
