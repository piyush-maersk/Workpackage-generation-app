"""
Workpackage Generation App
==========================
Prototype v0.1 – Automation & Operational Technology

RAG-powered automated workpackage document generator.
Takes a scope-of-work text and/or a FbM Tech Estimation Excel sheet as input
and produces a populated DOCX workpackage for the selected template type.

Run:
    streamlit run app.py
"""
from __future__ import annotations

import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ── Template registry ─────────────────────────────────────────────────────────

TEMPLATES: dict[str, dict] = {
    "IT Device": {
        "file": "IT Hardware - Workpackage Template.docx",
        "available": True,
        "category": "Hardware",
        "icon": "💻",
        "description": "Laptops, desktops, monitors, peripherals, meeting-room equipment",
    },
    "OT Device": {
        "file": "OT Hardware - Workpackage Template.docx",
        "available": True,
        "category": "Hardware",
        "icon": "📦",
        "description": "RF scanners, label printers, mobile printers, tablets, battery chargers",
    },
    "MDF": {
        "file": "MDF_Workpackage.docx",
        "available": True,
        "category": "Network",
        "icon": "🔌",
        "description": "Main Distribution Frame – core network infrastructure and racks",
    },
    "OT Automation Machine": {
        "file": "OT_Automationmachine_Workpackage.docx",
        "available": True,
        "category": "Automation",
        "icon": "🤖",
        "description": "Cubiscan, RFID tunnels, PANDA, carton erectors, automation machines",
    },
    "Perimeter": {
        "file": None,
        "available": False,
        "category": "Network",
        "icon": "🔒",
        "description": "Firewall, edge routers, WAN security",
    },
    "IT Net": {
        "file": None,
        "available": False,
        "category": "Network",
        "icon": "🌐",
        "description": "Access switches, access points, structured cabling",
    },
    "OT Net": {
        "file": None,
        "available": False,
        "category": "Network",
        "icon": "📡",
        "description": "OT network components and segmentation",
    },
    "Automation": {
        "file": None,
        "available": False,
        "category": "Automation",
        "icon": "⚙️",
        "description": "General automation workpackage",
    },
}


# ── App entry point ───────────────────────────────────────────────────────────

def main() -> None:
    st.set_page_config(
        page_title="Workpackage Generator",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("📋  Workpackage Generation App")
    st.markdown(
        "*Automation & Operational Technology – Automated Document Generation · Prototype v0.1*"
    )
    st.divider()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️  Configuration")
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            value=os.getenv("OPENAI_API_KEY", ""),
            help="Required for AI-powered extraction and content generation",
        )
        model = st.selectbox(
            "LLM Model",
            ["gpt-4o", "gpt-4-turbo", "gpt-4", "gpt-3.5-turbo"],
            index=0,
            help="GPT-4o / GPT-4 recommended for best accuracy",
        )

        st.divider()
        st.markdown("**Template availability**")
        for name, info in TEMPLATES.items():
            badge = "✅" if info["available"] else "🔜"
            st.markdown(f"{badge} {info['icon']} **{name}** _{info['category']}_")

        st.divider()
        st.caption("Prototype v0.1 · Maersk AOT")

    # ── Step 1: Template selection ────────────────────────────────────────────
    st.header("1️⃣  Select Workpackage Template")

    template_cols = st.columns(4)
    for idx, (name, info) in enumerate(TEMPLATES.items()):
        with template_cols[idx % 4]:
            if info["available"]:
                st.markdown(
                    f"**{info['icon']} {name}**  \n"
                    f"<span style='color:green'>Available</span>  \n"
                    f"_{info['category']}_",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"~~{info['icon']} {name}~~  \n"
                    f"<span style='color:grey'>Yet to Arrive</span>  \n"
                    f"_{info['category']}_",
                    unsafe_allow_html=True,
                )

    selected_name = st.selectbox(
        "Choose a template:",
        options=list(TEMPLATES.keys()),
        format_func=lambda x: (
            f"{'✅' if TEMPLATES[x]['available'] else '🔜'} "
            f"{TEMPLATES[x]['icon']} {x}"
            + ("" if TEMPLATES[x]["available"] else "  [Yet to Arrive]")
        ),
    )
    selected = TEMPLATES[selected_name]

    if not selected["available"]:
        st.warning(
            f"⏳ **{selected_name}** is not yet available.  "
            "Please select one of the 4 available templates."
        )
        st.stop()

    st.success(
        f"Selected: **{selected['icon']} {selected_name}** — {selected['description']}"
    )
    st.divider()

    # ── Step 2: Project information ───────────────────────────────────────────
    st.header("2️⃣  Project Information")
    st.caption(
        "Fill in what you know – remaining fields will be auto-extracted from your documents."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        project_name = st.text_input(
            "Project Name", placeholder="e.g., Customer Integration – Warehouse A"
        )
        fbm_id = st.text_input("FBM ID", placeholder="e.g., FBM-2026-001")
    with c2:
        site_id = st.text_input("Site ID", placeholder="e.g., SITE-001")
        customer_name = st.text_input("Customer Name", placeholder="e.g., Acme Corp")
    with c3:
        site_address = st.text_input(
            "Site Address", placeholder="e.g., 123 Warehouse Blvd, Amsterdam"
        )
        go_live_date = st.text_input("Go-Live Date", placeholder="e.g., 01/06/2026")

    st.divider()

    # ── Step 3: Input data ────────────────────────────────────────────────────
    st.header("3️⃣  Provide Input Data")

    tab_scope, tab_est = st.tabs([
        "📝 Scope of Work",
        "📊 Cost Estimation Sheet (Excel / CSV)",
    ])

    with tab_scope:
        scope_text = st.text_area(
            "Paste your Scope of Work / Project Description",
            height=260,
            placeholder=(
                "Enter the project description and device list here…\n\n"
                "Example:\n"
                "The project is a Customer Integration for warehouse operation "
                "with total area of X sq ft. Scope includes provisioning laptops, "
                "desktops, barcode scanners, label printers, RF scanners, and "
                "network components.\n\n"
                "Laptop Computer: 8\n"
                "Desktop Computer (Packing Station): 29\n"
                "Standard Radio Frequency Scanner: 20\n"
                "Industrial Label Printer (Packing Station): 29\n"
                "…"
            ),
        )
        scope_file = st.file_uploader(
            "Or upload a .txt file with the scope of work",
            type=["txt"],
            key="scope_file",
        )

    with tab_est:
        st.markdown(
            "Upload the **FbM Tech Estimation** Excel sheet.  \n"
            "The tool will automatically extract device descriptions and quantities.  \n"
            "Supported formats: `.xlsx`, `.xls`, `.csv`"
        )
        estimation_file = st.file_uploader(
            "Upload Cost Estimation Sheet",
            type=["csv", "xlsx", "xls"],
            help="Supports the FbM Tech Estimation layout with Description and Quantity columns",
        )
        if estimation_file:
            st.info(f"📁  Uploaded: `{estimation_file.name}`")

    st.divider()

    # ── Generate ──────────────────────────────────────────────────────────────
    if not api_key:
        st.warning("⚠️  Enter your **OpenAI API Key** in the sidebar to enable generation.")

    can_generate = bool(api_key) and (
        bool(scope_text) or bool(scope_file) or bool(estimation_file)
    )

    if st.button(
        "🚀  Generate Workpackage Document",
        type="primary",
        disabled=not can_generate,
        use_container_width=True,
    ):
        _run_generation(
            template_name=selected_name,
            template_info=selected,
            project_info={
                "project_name": project_name,
                "fbm_id": fbm_id,
                "site_id": site_id,
                "site_address": site_address,
                "customer_name": customer_name,
                "go_live_date": go_live_date,
            },
            scope_text=scope_text,
            scope_file=scope_file,
            estimation_file=estimation_file,
            api_key=api_key,
            model=model,
        )


# ── Generation pipeline ───────────────────────────────────────────────────────

def _run_generation(
    template_name: str,
    template_info: dict,
    project_info: dict[str, str],
    scope_text: str,
    scope_file: object,
    estimation_file: object,
    api_key: str,
    model: str,
) -> None:
    """Orchestrate the full RAG workpackage generation pipeline."""
    # Lazy imports – only loaded when generation is triggered
    from src.parser import InputParser
    from src.device_classifier import DeviceClassifier
    from src.rag_engine import RAGEngine
    from src.template_filler import TemplateFiller

    progress = st.progress(0, "Initialising…")
    status = st.empty()

    try:
        # ── Initialise components ──────────────────────────────────────────────
        status.info("🔧  Initialising RAG engine…")
        engine = RAGEngine(api_key=api_key, model=model)
        parser = InputParser()
        classifier = DeviceClassifier()
        filler = TemplateFiller()
        progress.progress(8)

        # ── Index templates ────────────────────────────────────────────────────
        status.info("📚  Indexing template knowledge base…")
        template_dir = Path(__file__).parent
        engine.index_templates(template_dir)
        progress.progress(22)

        # ── Parse inputs ───────────────────────────────────────────────────────
        status.info("📖  Parsing input documents…")
        all_devices: list[dict] = []
        combined_scope = ""

        if scope_file:
            combined_scope = scope_file.read().decode("utf-8", errors="replace")
        elif scope_text:
            combined_scope = scope_text

        if combined_scope:
            parsed = parser.parse_scope_text(combined_scope)
            all_devices.extend(parsed.get("devices", []))

        if estimation_file:
            suffix = Path(estimation_file.name).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(estimation_file.read())
                tmp_path = tmp.name
            try:
                est_parsed = parser.parse_estimation_file(tmp_path)
                all_devices.extend(est_parsed.get("devices", []))
            finally:
                os.unlink(tmp_path)

        # Deduplicate – keep highest quantity for each device name
        seen: dict[str, dict] = {}
        for d in all_devices:
            key = d["name"].strip().lower()
            if key not in seen or d["quantity"] > seen[key]["quantity"]:
                seen[key] = d
        all_devices = list(seen.values())
        progress.progress(38)

        with st.expander(
            f"📦  Extracted {len(all_devices)} device type(s) from input", expanded=False
        ):
            for d in all_devices:
                st.write(f"- **{d['name']}** — qty {d['quantity']}")

        # ── Classify devices ───────────────────────────────────────────────────
        status.info("🏷️  Classifying devices (IT / OT / Network / MDF / Automation)…")
        classified = classifier.classify_all(all_devices)
        progress.progress(50)

        with st.expander("🔍  Device Classification", expanded=False):
            for cat, devs in classified.items():
                if devs:
                    st.markdown(f"**{cat}** ({len(devs)} type(s))")
                    for d in devs:
                        st.write(f"  · {d['name']}: {d['quantity']}")

        # ── Extract project info ───────────────────────────────────────────────
        status.info("📋  Extracting project information with LLM…")
        extracted = engine.extract_project_info(
            scope_text=combined_scope,
            devices=all_devices,
        )
        # User-provided values override auto-extracted ones
        for key, val in project_info.items():
            if val:
                extracted[key] = val
            elif key not in extracted or not extracted.get(key):
                extracted[key] = ""
        progress.progress(64)

        # ── Generate content ───────────────────────────────────────────────────
        status.info(f"✍️  Generating {template_name} workpackage content with RAG…")
        generated = engine.generate_workpackage_content(
            template_type=template_name,
            project_info=extracted,
            classified_devices=classified,
            scope_text=combined_scope,
        )
        progress.progress(82)

        # ── Fill template ──────────────────────────────────────────────────────
        status.info("📄  Populating DOCX template…")
        template_path = Path(__file__).parent / template_info["file"]
        output_doc = filler.fill_template(
            template_path=str(template_path),
            template_type=template_name,
            project_info=extracted,
            classified_devices=classified,
            generated_content=generated,
        )

        buf = io.BytesIO()
        output_doc.save(buf)
        buf.seek(0)
        progress.progress(100)
        status.success("✅  Workpackage generated successfully!")

        # ── Summary ────────────────────────────────────────────────────────────
        with st.expander("📄  Generation Summary", expanded=True):
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(
                "IT Devices",
                sum(d["quantity"] for d in classified.get("IT", [])),
            )
            m2.metric(
                "OT Devices",
                sum(d["quantity"] for d in classified.get("OT", [])),
            )
            m3.metric(
                "Network Items",
                sum(d["quantity"] for d in classified.get("Network", [])),
            )
            m4.metric(
                "Software / Other",
                sum(d["quantity"] for d in classified.get("Software", []))
                + sum(d["quantity"] for d in classified.get("Other", [])),
            )

            if generated.get("summary"):
                st.markdown("**AI-Generated Summary:**")
                st.info(generated["summary"])

        # ── Download ───────────────────────────────────────────────────────────
        proj = extracted.get("project_name") or "Workpackage"
        fname = (
            f"{template_name.replace(' ', '_')}_"
            f"{proj.replace(' ', '_')}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )
        st.download_button(
            label="⬇️  Download Workpackage (.docx)",
            data=buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )

    except (ValueError, KeyError, FileNotFoundError) as exc:
        progress.empty()
        st.error(f"❌  Configuration or file error: {exc}")
    except Exception as exc:  # noqa: BLE001 – catch-all to keep prototype UI stable
        progress.empty()
        import traceback

        st.error(f"❌  Unexpected error: {exc}")
        with st.expander("🐛  Full traceback"):
            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
