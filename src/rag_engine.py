"""
RAG Engine – uses LangChain + OpenAI + ChromaDB to:

  1. Index the four available workpackage template DOCX files as a vector
     knowledge base (ChromaDB, in-memory for prototype).
  2. Extract structured project metadata from user-supplied scope text.
  3. Generate workpackage section content using retrieved template context.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from langchain_community.document_loaders import Docx2txtLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings, ChatOpenAI
from langchain_community.vectorstores import Chroma


class RAGEngine:
    """Retrieval-Augmented Generation engine for workpackage content."""

    TEMPLATE_FILES: dict[str, str] = {
        "IT Device": "IT Hardware - Workpackage Template.docx",
        "OT Device": "OT Hardware - Workpackage Template.docx",
        "MDF": "MDF_Workpackage.docx",
        "OT Automation Machine": "OT_Automationmachine_Workpackage.docx",
    }

    def __init__(self, api_key: str, model: str = "gpt-4o") -> None:
        self.api_key = api_key
        self.model = model
        self._vectorstore: Chroma | None = None
        self._llm = ChatOpenAI(
            api_key=api_key,
            model=model,
            temperature=0.2,
        )
        self._embeddings = OpenAIEmbeddings(api_key=api_key)

    # ── Template indexing ──────────────────────────────────────────────────────

    def index_templates(self, template_dir: Path) -> None:
        """Load and index all available DOCX templates into ChromaDB."""
        docs = []
        splitter = RecursiveCharacterTextSplitter(
            chunk_size=800,
            chunk_overlap=100,
        )
        for tname, fname in self.TEMPLATE_FILES.items():
            fpath = template_dir / fname
            if not fpath.exists():
                continue
            try:
                loader = Docx2txtLoader(str(fpath))
                raw = loader.load()
                for d in raw:
                    d.metadata["template_type"] = tname
                    d.metadata["source_file"] = fname
                chunks = splitter.split_documents(raw)
                docs.extend(chunks)
            except Exception:  # noqa: BLE001
                continue

        if docs:
            self._vectorstore = Chroma.from_documents(
                docs,
                self._embeddings,
                collection_name="workpackage_templates",
            )

    # ── Project info extraction ────────────────────────────────────────────────

    def extract_project_info(
        self,
        scope_text: str,
        devices: list[dict],
    ) -> dict[str, str]:
        """
        Use the LLM to extract structured project metadata from scope text.

        Returns a dict with keys: project_name, customer_name, site_address,
        site_id, fbm_id, go_live_date, access_date, description.
        """
        if not scope_text.strip():
            return {}

        device_summary = ", ".join(
            f"{d['name']} ×{d['quantity']}" for d in devices[:10]
        )

        prompt = (
            "You are an expert project analyst. "
            "Extract structured project information from the scope-of-work text below.\n\n"
            f"Scope of Work:\n\"\"\"\n{scope_text[:3000]}\n\"\"\"\n\n"
            f"Device summary (for context): {device_summary}\n\n"
            "Return a JSON object with these fields (use empty string if unknown):\n"
            "{\n"
            '  "project_name": "",\n'
            '  "customer_name": "",\n'
            '  "site_address": "",\n'
            '  "site_id": "",\n'
            '  "fbm_id": "",\n'
            '  "go_live_date": "",\n'
            '  "access_date": "",\n'
            '  "description": ""\n'
            "}\n\n"
            "Return ONLY valid JSON, no extra text."
        )
        return self._invoke_json(prompt, default={
            "project_name": "",
            "customer_name": "",
            "site_address": "",
            "site_id": "",
            "fbm_id": "",
            "go_live_date": "",
            "access_date": "",
            "description": "",
        })

    # ── Workpackage content generation ────────────────────────────────────────

    def generate_workpackage_content(
        self,
        template_type: str,
        project_info: dict[str, str],
        classified_devices: dict[str, list[dict]],
        scope_text: str = "",
    ) -> dict[str, Any]:
        """
        Generate section content for the selected template using RAG.

        Retrieves relevant template chunks from ChromaDB and passes them as
        grounding context to the LLM for generation.
        """
        context = self._retrieve_context(template_type, classified_devices)

        generators = {
            "IT Device": self._gen_it_device,
            "OT Device": self._gen_ot_device,
            "MDF": self._gen_mdf,
            "OT Automation Machine": self._gen_ot_automation,
        }
        gen_fn = generators.get(template_type, self._gen_generic)
        return gen_fn(
            project_info=project_info,
            classified_devices=classified_devices,
            scope_text=scope_text,
            context=context,
        )

    # ── Per-template generators ───────────────────────────────────────────────

    def _gen_it_device(
        self,
        project_info: dict,
        classified_devices: dict,
        scope_text: str,
        context: str,
    ) -> dict[str, Any]:
        it_devices = classified_devices.get("IT", [])
        sw_devices = classified_devices.get("Software", [])

        device_list = "\n".join(
            f"  - {d['name']}: {d['quantity']} unit(s)" for d in it_devices
        ) or "  (none specified)"
        sw_list = "\n".join(
            f"  - {d['name']}: {d['quantity']} unit(s)" for d in sw_devices
        ) or "  (none specified)"

        # Derive station counts from device names + quantities
        admin_qty = sum(
            d["quantity"] for d in it_devices
            if any(kw in d["name"].lower() for kw in ("laptop", "notebook", "elitebook"))
        )
        pack_qty = sum(
            d["quantity"] for d in it_devices
            if "packing" in d["name"].lower() or "pack station" in d["name"].lower()
        )

        prompt = (
            "You are a professional technical writer creating an IT Hardware Work Package "
            "for a warehouse project.\n\n"
            f"Project: {project_info.get('project_name', 'N/A')}\n"
            f"Customer: {project_info.get('customer_name', 'N/A')}\n"
            f"Site: {project_info.get('site_address', 'N/A')}\n\n"
            f"IT Hardware identified:\n{device_list}\n\n"
            f"Software/Licenses:\n{sw_list}\n\n"
            f"Template reference (RAG context):\n{context[:1500]}\n\n"
            "Generate the following in JSON format:\n"
            "{\n"
            '  "summary": "2-3 sentence executive summary of IT hardware scope",\n'
            '  "station_summary": "Brief description of user and workstation types",\n'
            '  "bom_notes": "Any notes about the Bill of Materials",\n'
            f'  "admin_users": {admin_qty or 0},\n'
            f'  "pack_stations": {pack_qty or 0},\n'
            '  "inbound_stations": 0,\n'
            '  "return_stations": 0,\n'
            '  "other_stations": 0,\n'
            '  "scope_narrative": "One paragraph scope of work narrative"\n'
            "}\n\n"
            "Return ONLY valid JSON."
        )
        return self._invoke_json(prompt, default={
            "summary": "",
            "station_summary": "",
            "bom_notes": "",
            "admin_users": admin_qty,
            "pack_stations": pack_qty,
            "inbound_stations": 0,
            "return_stations": 0,
            "other_stations": 0,
            "scope_narrative": "",
        })

    def _gen_ot_device(
        self,
        project_info: dict,
        classified_devices: dict,
        scope_text: str,
        context: str,
    ) -> dict[str, Any]:
        ot_devices = classified_devices.get("OT", [])
        device_list = "\n".join(
            f"  - {d['name']}: {d['quantity']} unit(s)" for d in ot_devices
        ) or "  (none specified)"

        prompt = (
            "You are a professional technical writer creating an OT Hardware Work Package "
            "for a warehouse project.\n\n"
            f"Project: {project_info.get('project_name', 'N/A')}\n"
            f"Site: {project_info.get('site_address', 'N/A')}\n\n"
            f"OT Hardware identified:\n{device_list}\n\n"
            f"Template reference (RAG context):\n{context[:1500]}\n\n"
            "Generate the following in JSON format:\n"
            "{\n"
            '  "summary": "2-3 sentence executive summary of OT hardware scope",\n'
            '  "rf_scanner_notes": "Notes about RF scanner configuration requirements",\n'
            '  "printer_notes": "Notes about label printer setup",\n'
            '  "soti_notes": "Notes about SOTI MDM enrollment",\n'
            '  "scope_narrative": "One paragraph scope of work narrative"\n'
            "}\n\n"
            "Return ONLY valid JSON."
        )
        return self._invoke_json(prompt, default={
            "summary": "",
            "rf_scanner_notes": "",
            "printer_notes": "",
            "soti_notes": "",
            "scope_narrative": "",
        })

    def _gen_mdf(
        self,
        project_info: dict,
        classified_devices: dict,
        scope_text: str,
        context: str,
    ) -> dict[str, Any]:
        network = classified_devices.get("Network", [])
        mdf = classified_devices.get("MDF", [])
        all_net = network + mdf

        device_list = "\n".join(
            f"  - {d['name']}: {d['quantity']} unit(s)" for d in all_net
        ) or "  (none specified)"

        # Estimate user count from IT devices if available
        it_devs = classified_devices.get("IT", [])
        users = sum(d["quantity"] for d in it_devs
                    if any(k in d["name"].lower() for k in ("laptop", "desktop", "computer")))

        prompt = (
            "You are a professional technical writer creating an MDF (Main Distribution Frame) "
            "Work Package for a warehouse project.\n\n"
            f"Project: {project_info.get('project_name', 'N/A')}\n"
            f"Site: {project_info.get('site_address', 'N/A')}\n\n"
            f"Network / MDF Hardware identified:\n{device_list}\n\n"
            f"Template reference (RAG context):\n{context[:1500]}\n\n"
            "Generate the following in JSON format:\n"
            "{\n"
            '  "summary": "2-3 sentence executive summary of MDF/network scope",\n'
            f'  "idf_count": 1,\n'
            f'  "rack_count": 1,\n'
            f'  "users_supported": {users or 0},\n'
            '  "scope_narrative": "One paragraph MDF scope narrative"\n'
            "}\n\n"
            "Return ONLY valid JSON."
        )
        return self._invoke_json(prompt, default={
            "summary": "",
            "idf_count": 1,
            "rack_count": 1,
            "users_supported": users,
            "scope_narrative": "",
        })

    def _gen_ot_automation(
        self,
        project_info: dict,
        classified_devices: dict,
        scope_text: str,
        context: str,
    ) -> dict[str, Any]:
        auto_devices = classified_devices.get("Automation", [])
        device_list = "\n".join(
            f"  - {d['name']}: {d['quantity']} unit(s)" for d in auto_devices
        ) or "  (standard automation scope applies – no specific devices listed)"

        prompt = (
            "You are a professional technical writer creating an OT Automation Machine "
            "Work Package for a warehouse project.\n\n"
            f"Project: {project_info.get('project_name', 'N/A')}\n"
            f"Customer: {project_info.get('customer_name', 'N/A')}\n"
            f"Site: {project_info.get('site_address', 'N/A')}\n\n"
            f"Automation Equipment identified:\n{device_list}\n\n"
            f"Template reference (RAG context):\n{context[:1500]}\n\n"
            "Generate the following in JSON format:\n"
            "{\n"
            '  "summary": "2-3 sentence executive summary of OT automation scope",\n'
            '  "equipment_narrative": "Paragraph describing the automation equipment scope",\n'
            '  "integration_notes": "Notes about WMS/TMS integration requirements",\n'
            '  "scope_narrative": "One paragraph scope of work narrative"\n'
            "}\n\n"
            "Return ONLY valid JSON."
        )
        return self._invoke_json(prompt, default={
            "summary": "",
            "equipment_narrative": "",
            "integration_notes": "",
            "scope_narrative": "",
        })

    def _gen_generic(self, **_kwargs: Any) -> dict[str, Any]:
        return {"summary": "", "scope_narrative": ""}

    # ── Retrieval helpers ──────────────────────────────────────────────────────

    def _retrieve_context(
        self,
        template_type: str,
        classified_devices: dict,
    ) -> str:
        """Retrieve relevant template chunks from ChromaDB for grounding."""
        if self._vectorstore is None:
            return ""

        # Build query from template type + representative device names
        terms: list[str] = [template_type]
        for cat_devices in classified_devices.values():
            for d in cat_devices[:3]:
                terms.append(d["name"])
        query = " ".join(terms[:8])

        try:
            # Try filtered retrieval first
            retriever = self._vectorstore.as_retriever(
                search_kwargs={
                    "k": 5,
                    "filter": {"template_type": template_type},
                }
            )
            docs = retriever.invoke(query)
            if not docs:
                # Fall back to unfiltered
                retriever = self._vectorstore.as_retriever(search_kwargs={"k": 5})
                docs = retriever.invoke(query)
            return "\n\n---\n\n".join(d.page_content for d in docs)
        except Exception:  # noqa: BLE001
            return ""

    # ── LLM utility ───────────────────────────────────────────────────────────

    def _invoke_json(self, prompt: str, default: dict) -> dict:
        """Invoke the LLM and parse the JSON response; return default on failure."""
        try:
            response = self._llm.invoke(prompt)
            text = response.content.strip()
            # Strip Markdown code fences if present
            text = re.sub(r"^```(?:json)?\s*", "", text)
            text = re.sub(r"\s*```$", "", text)
            return json.loads(text)
        except Exception:  # noqa: BLE001
            return default
