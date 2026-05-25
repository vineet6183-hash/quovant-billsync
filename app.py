import os
import sys
import zipfile
import tempfile
import shutil
from datetime import datetime

import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(__file__))
import pdf_parser
import billsync_formatter
import tracker_manager

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QUOVANT BillSync",
    page_icon="⚖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state ─────────────────────────────────────────────────────────────
_defaults = {
    "results": [],
    "all_items_combined": [],
    "processing_done": False,
    "billsync_bytes": None,
    "tracker_bytes": None,
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ── Helpers ───────────────────────────────────────────────────────────────────

def _save_to_temp(uploaded_file) -> str:
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


def _pdfs_from_zip(zip_path: str) -> list:
    """Returns [(display_name, extracted_path), ...]"""
    temp_dir = tempfile.mkdtemp()
    found = []
    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.namelist():
            if member.lower().endswith(".pdf"):
                dest = zf.extract(member, path=temp_dir)
                # flatten name to avoid subdirectory collisions in display
                display = member.replace("/", "_").replace("\\", "_")
                found.append((display, dest))
    return found, temp_dir


def _build_billsync_bytes(all_items: list) -> bytes | None:
    if not all_items:
        return None
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name
    try:
        billsync_formatter.format_to_billsync(all_items, None, tmp_path)
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _build_tracker_bytes(results: list) -> bytes | None:
    good = [r for r in results if r["error"] is None]
    if not good:
        return None
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tracker_path = tmp.name
    # remove file so tracker_manager creates it fresh
    os.unlink(tracker_path)
    try:
        for r in good:
            tracker_manager.update_tracker(tracker_path, r["metadata"], r["filename"])
        with open(tracker_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tracker_path):
            os.unlink(tracker_path)


def process_files(uploaded_files, update_tracker_flag: bool):
    results = []
    all_items = []
    temp_dirs_to_clean = []
    temp_files_to_clean = []

    progress = st.progress(0, text="Starting…")
    total = len(uploaded_files)

    try:
        for idx, uf in enumerate(uploaded_files):
            progress.progress(idx / total, text=f"Processing {uf.name}…")

            pdfs_to_parse = []  # [(display_name, path)]

            if uf.name.lower().endswith(".zip"):
                zip_tmp = _save_to_temp(uf)
                temp_files_to_clean.append(zip_tmp)
                try:
                    extracted, tmp_dir = _pdfs_from_zip(zip_tmp)
                    temp_dirs_to_clean.append(tmp_dir)
                    if not extracted:
                        results.append({
                            "filename": uf.name,
                            "items": [],
                            "metadata": {},
                            "matter_number": "",
                            "error": "No PDF files found inside ZIP.",
                        })
                        continue
                    pdfs_to_parse.extend(extracted)
                except Exception as e:
                    results.append({
                        "filename": uf.name,
                        "items": [],
                        "metadata": {},
                        "matter_number": "",
                        "error": f"ZIP extraction failed: {e}",
                    })
                    continue
            else:
                pdf_tmp = _save_to_temp(uf)
                temp_files_to_clean.append(pdf_tmp)
                pdfs_to_parse.append((uf.name, pdf_tmp))

            for pdf_name, pdf_path in pdfs_to_parse:
                try:
                    items, metadata = pdf_parser.extract_invoice_data(pdf_path)
                    matter_number = items[0].get("matter_number", "") if items else ""
                    all_items.extend(items)
                    results.append({
                        "filename": pdf_name,
                        "items": items,
                        "metadata": metadata,
                        "matter_number": matter_number,
                        "error": None,
                    })
                except Exception as e:
                    results.append({
                        "filename": pdf_name,
                        "items": [],
                        "metadata": {},
                        "matter_number": "",
                        "error": str(e),
                    })

        progress.progress(1.0, text="Done.")

    finally:
        for p in temp_files_to_clean:
            try:
                os.unlink(p)
            except Exception:
                pass
        for d in temp_dirs_to_clean:
            try:
                shutil.rmtree(d, ignore_errors=True)
            except Exception:
                pass

    st.session_state.results = results
    st.session_state.all_items_combined = all_items
    st.session_state.processing_done = True
    st.session_state.billsync_bytes = _build_billsync_bytes(all_items)
    st.session_state.tracker_bytes = (
        _build_tracker_bytes(results) if update_tracker_flag else None
    )


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## QUOVANT")
    st.markdown("**BillSync Automation**  \n*Web Edition*")
    st.divider()

    uploaded_files = st.file_uploader(
        "Upload Invoice Files",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        help="Upload one or more PDF invoice files or ZIP archives containing PDFs.",
    )

    st.divider()

    update_tracker_flag = st.checkbox(
        "Update Master Tracker",
        value=False,
        help=(
            "Appends each processed invoice to a Master_Tracker.xlsx available for download. "
            "Note: hyperlinks in the tracker will point to temporary paths — "
            "replace them manually if needed after downloading."
        ),
    )

    process_btn = st.button(
        "Process Files",
        type="primary",
        disabled=not uploaded_files,
        use_container_width=True,
    )

    if process_btn and uploaded_files:
        for k, v in _defaults.items():
            st.session_state[k] = v
        process_files(uploaded_files, update_tracker_flag)
        st.rerun()

    if st.session_state.processing_done:
        st.divider()
        st.markdown("**Downloads**")

        if st.session_state.billsync_bytes:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="Download BillSync Excel",
                data=st.session_state.billsync_bytes,
                file_name=f"BillSync_Output_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.warning("No reduced line items found — BillSync output is empty.")

        if st.session_state.tracker_bytes:
            st.download_button(
                label="Download Master Tracker",
                data=st.session_state.tracker_bytes,
                file_name="Master_Tracker.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ── Main area: welcome state ──────────────────────────────────────────────────
if not st.session_state.processing_done:
    st.title("QUOVANT BillSync Automation")
    st.markdown(
        """
        Upload PDF invoice files or ZIP archives from the sidebar, then click **Process Files**.

        **Supported inputs**
        - Individual `.pdf` files — Quovant compliance report invoices
        - `.zip` archives — one or more PDFs inside

        **What you get**
        - Live dashboard with per-invoice summaries (amounts, reductions, line items)
        - **BillSync Excel** download — 22-column format, reduced line items only
        - Optionally a **Master Tracker** Excel with one row per invoice
        """
    )
    st.stop()

# ── Main area: dashboard ──────────────────────────────────────────────────────
results = st.session_state.results
successful = [r for r in results if r["error"] is None]
failed = [r for r in results if r["error"] is not None]

st.title("Processing Results")
st.caption(
    f"{len(successful)} invoice(s) processed successfully"
    + (f" · {len(failed)} error(s)" if failed else "")
)

# Aggregate metrics
total_submitted = sum(r["metadata"].get("total_submitted", 0) for r in successful)
total_reduction = sum(r["metadata"].get("total_reduction", 0) for r in successful)
total_payable = sum(r["metadata"].get("total_payable", 0) for r in successful)
reduction_pct = (total_reduction / total_submitted * 100) if total_submitted else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Invoices Processed", len(successful))
c2.metric("Total Submitted", f"${total_submitted:,.2f}")
c3.metric("Total Reduction", f"${total_reduction:,.2f}", f"{reduction_pct:.1f}%")
c4.metric("Total Payable", f"${total_payable:,.2f}")

st.divider()

# Per-invoice expanders
st.subheader("Invoice Summary")

for r in results:
    meta = r["metadata"]
    inv_no = meta.get("invoice_number") or r["filename"]
    label = f"Invoice {inv_no}  ·  {r['filename']}"

    with st.expander(label, expanded=(len(results) <= 4)):
        if r["error"]:
            st.error(f"Parse error: {r['error']}")
            continue

        hc1, hc2, hc3, hc4 = st.columns(4)
        hc1.markdown(f"**Invoice #**  \n{meta.get('invoice_number', 'N/A')}")
        hc2.markdown(f"**Matter #**  \n{r.get('matter_number', 'N/A')}")
        hc3.markdown(f"**Invoice Date**  \n{meta.get('invoice_date', 'N/A')}")
        hc4.markdown(f"**Finalized Date**  \n{meta.get('finalized_date', 'N/A')}")

        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("Submitted", f"${meta.get('total_submitted', 0):,.2f}")
        mc2.metric("Reduction", f"${meta.get('total_reduction', 0):,.2f}")
        mc3.metric("Payable", f"${meta.get('total_payable', 0):,.2f}")

        n_total = len(r["items"])
        n_reduced = sum(1 for it in r["items"] if abs(it.get("reduced_amount", 0)) > 0.001)
        st.caption(f"{n_total} total line items · {n_reduced} with reductions")

# Errors
if failed:
    st.divider()
    st.subheader("Failed Files")
    for r in failed:
        st.error(f"**{r['filename']}**: {r['error']}")

# Line items table
st.divider()
st.subheader("Line Items")

all_items = st.session_state.all_items_combined

if all_items:
    df = pd.DataFrame(all_items)

    display_cols = [
        "invoice_number", "matter_number", "invoice_date",
        "line_no", "date", "timekeeper", "item_type",
        "rate", "units", "amount", "reduced_amount", "reason",
    ]
    display_cols = [c for c in display_cols if c in df.columns]
    df_display = df[display_cols].copy()
    df_display.rename(
        columns={
            "invoice_number": "Invoice #",
            "matter_number": "Matter #",
            "invoice_date": "Invoice Date",
            "line_no": "Line",
            "date": "Item Date",
            "timekeeper": "Timekeeper",
            "item_type": "Type",
            "rate": "Rate",
            "units": "Units",
            "amount": "Amount",
            "reduced_amount": "Reduced Amount",
            "reason": "Reduction Reason",
        },
        inplace=True,
    )

    invoice_options = ["All"] + sorted(
        df_display["Invoice #"].dropna().unique().tolist()
    )
    selected_invoice = st.selectbox("Filter by Invoice", options=invoice_options)

    if selected_invoice != "All":
        df_display = df_display[df_display["Invoice #"] == selected_invoice]

    def _highlight_reduced(row):
        if abs(row.get("Reduced Amount", 0)) > 0.001:
            return ["background-color: #fff3cd"] * len(row)
        return [""] * len(row)

    st.dataframe(
        df_display.style.apply(_highlight_reduced, axis=1),
        use_container_width=True,
        height=450,
        hide_index=True,
    )

    n_reduced_rows = int((df_display["Reduced Amount"].abs() > 0.001).sum())
    st.caption(
        f"Showing {len(df_display)} rows · "
        f"{n_reduced_rows} with reductions (highlighted in yellow)"
    )
else:
    st.info("No line items extracted.")
