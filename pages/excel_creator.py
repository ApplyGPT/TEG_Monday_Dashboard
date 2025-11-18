"""
Excel Creator
Builds the Development Package section of the workbook using a template.
"""

from __future__ import annotations

import os
import re
from io import BytesIO

import streamlit as st

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment
except Exception:  # pragma: no cover - fallback if dependency missing at runtime
    load_workbook = None
    Font = None
    Alignment = None


BASE_PRICE = 2325.00
OPTIONAL_PRICES = {
    "wash_dye": 1330.00,
    "design": 1330.00,
    "source": 1330.00,
    "treatment": 760.00,
}
TEMPLATE_FILENAME = "Copy of TEG 2025 WORKBOOK TEMPLATES.xlsx"
TARGET_SHEET = "DEVELOPMENT ONLY"
ROW_INDICES = [10, 12, 14, 16, 18]  # Rows reserved for style entries
STYLE_OPTIONS = ["DRESS", "JACKET", "T-SHIRT", "SKIRT", "LS SHIRT"]


@st.cache_data(show_spinner=False)
def get_template_path() -> str:
    """Return the absolute path to the Excel template."""
    base_dir = os.path.dirname(os.path.dirname(__file__))
    template_path = os.path.join(base_dir, "inputs", TEMPLATE_FILENAME)
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"Template '{TEMPLATE_FILENAME}' was not found in the inputs folder."
        )
    return template_path


def update_header_labels(ws, client_name: str) -> None:
    """Ensure headers and client info match the spec."""
    header_map = {
        "H9": "WASH/DYE",
        "I9": "DESIGN",
        "J9": "SOURCE",
        "K9": "TREATMENT",
        "L9": "TOTAL",
    }
    for cell, label in header_map.items():
        ws[cell].value = label

    ws["B3"].value = "TEGMADE, JUST FOR"
    client_display = (client_name or "").strip().upper()
    ws["J3"].value = client_display
    if Font is not None:
        ws["J3"].font = Font(
            color="00C9A57A",
            name="Schibsted Grotesk",
            size=48,
            bold=True,
        )
    if Alignment is not None:
        ws["J3"].alignment = Alignment(horizontal="left", vertical="center")

    if Alignment is not None:
        center_cells = ["L20", "P10", "P11", "P12", "P13", "P20"]
        for ref in center_cells:
            ws[ref].alignment = Alignment(horizontal="center", vertical="center")


def clear_style_rows(ws) -> None:
    """Blank out all reserved style rows (B–L across the five slots)."""
    for row_idx in ROW_INDICES:
        for col_idx in range(2, 13):  # Columns B through L
            ws.cell(row=row_idx, column=col_idx).value = None
    # Also clear totals that depend on these cells
    ws["F20"].value = None
    ws["L20"].value = None


def apply_development_package(
    ws,
    *,
    client_name: str,
    client_email: str,
    representative: str,
    style_entries: list[dict[str, float]],
) -> tuple[float, float]:
    """Write the inputs into the workbook and return totals."""
    # Header metadata
    ws["D6"].value = client_email.strip()
    ws["J6"].value = (representative or "").strip().upper()
    ws["B8"].value = "DEVELOPMENT PACKAGE"

    optional_cells = {
        "H": "wash_dye",
        "I": "design",
        "J": "source",
        "K": "treatment",
    }

    total_development = 0.0
    total_optional = 0.0

    for idx, row_idx in enumerate(ROW_INDICES):
        if idx < len(style_entries):
            entry = style_entries[idx]
            style_name = entry.get("name", "").strip() or "STYLE"
            complexity_pct = float(entry.get("complexity", 0.0))

            ws[f"B{row_idx}"].value = idx + 1
            ws[f"C{row_idx}"].value = style_name
            ws[f"D{row_idx}"].value = BASE_PRICE
            if complexity_pct == 0:
                ws[f"E{row_idx}"].value = None
            else:
                ws[f"E{row_idx}"].value = complexity_pct / 100.0
            ws[f"F{row_idx}"].value = f"=D{row_idx}*(1+E{row_idx})"

            # Optional add-ons per row
            row_options = entry.get("options", {}) or {}
            row_optional_sum = 0.0
            for col_letter, key in optional_cells.items():
                cell_ref = f"{col_letter}{row_idx}"
                if row_options.get(key):
                    price = OPTIONAL_PRICES[key]
                    ws[cell_ref].value = price
                    row_optional_sum += price
                else:
                    ws[cell_ref].value = None
            ws[f"L{row_idx}"].value = f"=SUM(H{row_idx}:K{row_idx})"

            total_development += BASE_PRICE * (1 + complexity_pct / 100.0)
            total_optional += row_optional_sum
        else:
            # Ensure unused rows stay blank
            for col_idx in range(2, 14):
                ws.cell(row=row_idx, column=col_idx).value = None

    # Totals section
    ws["F20"].value = "=SUM(F10:F19)"
    ws["L20"].value = "=SUM(L10:L19)"
    return total_development, total_optional


def build_workbook_bytes(
    *,
    client_name: str,
    client_email: str,
    representative: str,
    style_entries: list[dict[str, float]],
) -> tuple[bytes, float, float]:
    """Load the template, update it, and return bytes plus totals."""
    if load_workbook is None:
        raise RuntimeError(
            "openpyxl is not installed. Please add it to the environment first."
        )

    template_path = get_template_path()
    wb = load_workbook(template_path)
    if TARGET_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Worksheet '{TARGET_SHEET}' is missing from the template."
        )
    ws = wb[TARGET_SHEET]

    update_header_labels(ws, client_name)
    clear_style_rows(ws)
    total_dev, total_optional = apply_development_package(
        ws,
        client_name=client_name,
        client_email=client_email,
        representative=representative,
        style_entries=style_entries,
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read(), total_dev, total_optional


def main() -> None:
    st.set_page_config(
        page_title="Excel Creator",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.title("📊 Excel Creator")

    st.markdown(
        """
<style>
    [data-testid="stSidebarNav"] a[href*="quickbooks_form"] { display: none !important; }
    [data-testid="stSidebarNav"] a[href*="signnow_form"] { display: none !important; }
</style>
""",
        unsafe_allow_html=True,
    )

    st.caption(
        "Fill in the Development Package inputs and download a formatted workbook "
        "based on the official template."
    )

    client_name = st.text_input("Client Name", placeholder="Enter client name")

    col_a, col_b = st.columns(2)
    with col_a:
        client_email = st.text_input(
            "Client Email", placeholder="client@email.com"
        )
    with col_b:
        representative = st.text_input(
            "Representative", placeholder="Enter representative"
        )

    selected_styles = st.multiselect(
        "Styles",
        STYLE_OPTIONS,
        placeholder="Select one or more styles",
    )

    style_entries: list[dict[str, float]] = []
    if selected_styles:
        st.markdown("**Configure each selected style**")
        for style in selected_styles:
            complexity_value = st.number_input(
                f"{style} Complexity (%)",
                min_value=0,
                max_value=200,
                value=100,
                step=5,
                format="%d",
                key=f"complexity_{style}",
            )
            add_on_cols = st.columns(4)
            style_options = {
                "wash_dye": add_on_cols[0].checkbox(
                    "Wash/Dye ($1,330)", key=f"{style}_wash_dye", value=False
                ),
                "design": add_on_cols[1].checkbox(
                    "Design ($1,330)", key=f"{style}_design", value=False
                ),
                "source": add_on_cols[2].checkbox(
                    "Source ($1,330)", key=f"{style}_source", value=False
                ),
                "treatment": add_on_cols[3].checkbox(
                    "Treatment ($360)", key=f"{style}_treatment", value=False
                ),
            }
            style_entries.append(
                {"name": style, "complexity": complexity_value, "options": style_options}
            )
    else:
        st.info("Select at least one style to configure complexity.")

    if not style_entries:
        st.info("Select at least one style to enable the generator.")
        return

    try:
        excel_bytes, _, _ = build_workbook_bytes(
            client_name=client_name,
            client_email=client_email,
            representative=representative,
            style_entries=style_entries,
        )
    except FileNotFoundError as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # pragma: no cover - streamlit runtime feedback
        st.error(f"Failed to build workbook: {exc}")
        return

    safe_client = re.sub(r"[^A-Za-z0-9_-]+", "_", (client_name or "").strip()) or "client"
    download_name = f"development_package_{safe_client.lower()}.xlsx"

    st.success("Workbook is ready.")
    st.download_button(
        label="Generate Workbook",
        data=excel_bytes,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__":
    main()

