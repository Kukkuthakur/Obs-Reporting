"""
Streamlit UI entry point for the OB reporting app.
"""

import streamlit as st

# Fix the imports here!
from shared.fields import TEMPLATES, FIELDS
from shared.templates import TEMPLATE_BLOCKS, BLOCKS

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
        default_index = choices.index(default) if default in choices else 0
        st.selectbox(label, choices, index=default_index, key=key)
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
            # We use 2 columns for layout
            cols = st.columns(2)
            
            # Counter for column placement
            col_idx = 0 
            
            for field_id in block_fields:
                field = FIELD_BY_ID.get(field_id)
                if not field:
                    continue
                
                # Handle optional fields
                if field.get("optional"):
                    # Put the checkbox in the current column
                    with cols[col_idx % 2]:
                        include = st.checkbox(
                            f"Include {field['label']}", key=f"inc_{field_id}"
                        )
                    col_idx += 1
                    
                    if not include:
                        continue
                
                # Render the actual input field
                with cols[col_idx % 2]:
                    render_field(field)
                col_idx += 1

    st.markdown("---")
    if st.button("Generate report preview"):
        st.info("Template engine not yet wired. Fields captured above.")


if __name__ == "__main__":
    main()
