"""
Input parser – handles scope-of-work text and FbM Tech Estimation
Excel/CSV files.

Extracted output is a list of dicts:
    [{"name": "Laptop Computer", "quantity": 8}, ...]
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


class InputParser:
    """Parse scope-of-work text and FbM Tech Estimation spreadsheets."""

    # ── Scope-of-work text parsing ────────────────────────────────────────────

    def parse_scope_text(self, text: str) -> dict[str, Any]:
        """
        Extract project description and device list from plain text.

        Recognises lines such as:
          "Laptop Computer: 8"
          "Desktop Computer (Packing Station) – 29"
          "29  Desktop Computer (Packing Station)"
        """
        lines = text.strip().splitlines()
        devices: list[dict] = []
        description_lines: list[str] = []

        # Pattern A: "Device Name: qty"  /  "Device Name – qty"  (trailing number)
        trailing = re.compile(r"^(.+?)[\s:–\-—]+(\d+)\s*$", re.UNICODE)
        # Pattern B: leading number  "qty  Device Name"
        leading = re.compile(r"^(\d+)\s{2,}(.+)$")

        for raw in lines:
            line = raw.strip()
            if not line:
                continue

            # Strip leading list indices such as "1. " or "25. "
            clean = re.sub(r"^\d+\.\s*", "", line)

            m = trailing.match(clean)
            if m:
                name = m.group(1).strip(" -–—:")
                qty = int(m.group(2))
                if qty > 0 and len(name) > 2:
                    devices.append({"name": name, "quantity": qty})
                    continue

            m2 = leading.match(clean)
            if m2:
                qty = int(m2.group(1))
                name = m2.group(2).strip()
                if qty > 0 and len(name) > 2:
                    devices.append({"name": name, "quantity": qty})
                    continue

            # Treat as description prose
            if len(clean) > 15:
                description_lines.append(clean)

        return {
            "description": " ".join(description_lines[:5]),
            "devices": devices,
        }

    # ── Estimation file parsing ───────────────────────────────────────────────

    def parse_estimation_file(self, file_path: str) -> dict[str, Any]:
        """
        Parse a FbM Tech Estimation Excel or CSV file.

        Returns {"devices": [{"name": ..., "quantity": ..., "section": ...}]}
        """
        suffix = Path(file_path).suffix.lower()
        try:
            if suffix in (".xlsx", ".xls"):
                return self._parse_excel(file_path)
            return self._parse_csv(file_path)
        except Exception as exc:  # noqa: BLE001
            return {"devices": [], "error": str(exc)}

    def _parse_excel(self, file_path: str) -> dict[str, Any]:
        """Try each sheet; return devices from the first sheet that yields results."""
        xl = pd.ExcelFile(file_path)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet, header=None)
            devices = self._extract_from_dataframe(df)
            if devices:
                return {"devices": devices, "sheet": sheet}
        return {"devices": []}

    def _parse_csv(self, file_path: str) -> dict[str, Any]:
        df = pd.read_csv(file_path, header=None)
        return {"devices": self._extract_from_dataframe(df)}

    # ── DataFrame extraction ──────────────────────────────────────────────────

    def _extract_from_dataframe(self, df: pd.DataFrame) -> list[dict]:
        """
        Heuristically locate Description and Quantity columns then extract rows
        where Quantity > 0.

        Strategy:
          1. Find a header row with recognisable column names.
          2. Fall back to scanning every row for a (text, positive-integer) pair.
        """
        devices: list[dict] = []

        # ── Strategy 1: look for a header row ─────────────────────────────────
        header_row_idx = desc_col_idx = qty_col_idx = None
        for row_idx in range(min(25, len(df))):
            vals = [str(v).strip().lower() for v in df.iloc[row_idx].fillna("")]
            desc_hits = [
                i for i, v in enumerate(vals)
                if any(kw in v for kw in ("description", "item", "device", "hardware", "product"))
            ]
            qty_hits = [
                i for i, v in enumerate(vals)
                if any(kw in v for kw in ("qty", "quantity", "count"))
                and "unit" not in v
            ]
            if desc_hits and qty_hits:
                header_row_idx = row_idx
                desc_col_idx = desc_hits[0]
                qty_col_idx = qty_hits[0]
                break

        if header_row_idx is not None:
            current_section = ""
            for row_idx in range(header_row_idx + 1, len(df)):
                row = df.iloc[row_idx]
                desc = str(row.iloc[desc_col_idx]).strip()
                qty_raw = row.iloc[qty_col_idx]

                # Detect section header rows (non-numeric Description, empty qty)
                if desc and not self._is_numeric(str(qty_raw)) and len(desc) > 2:
                    if not self._looks_like_device(desc):
                        current_section = desc
                        continue

                if self._is_valid_device(desc, qty_raw):
                    devices.append({
                        "name": desc,
                        "quantity": int(float(qty_raw)),
                        "section": current_section,
                    })
            if devices:
                return devices

        # ── Strategy 2: FbM format – scan rows for Description column ─────────
        # FbM format has: ISA Level | Section | CAPEX code | Description | … | Qty | …
        # Description is usually in column index 3 or 4; Quantity in column 6 or 7.
        devices = self._parse_fbm_format(df)
        if devices:
            return devices

        # ── Strategy 3: generic row scan ──────────────────────────────────────
        seen: set[str] = set()
        for row_idx in range(len(df)):
            row = df.iloc[row_idx].fillna("")
            str_vals = [str(v).strip() for v in row]

            desc = ""
            qty = 0
            for i, v in enumerate(str_vals):
                if (
                    len(v) > 3
                    and not self._is_numeric(v)
                    and not v.upper().startswith("CAPEX")
                    and not v.upper().startswith("OPEX")
                    and v.lower() not in ("local", "global", "n/a", "nan", "")
                    and not v.isupper()
                    and desc == ""
                ):
                    desc = v
                if self._is_numeric(v) and int(float(v)) > 0 and qty == 0 and i > 0:
                    qty = int(float(v))

            key = desc.lower()
            if desc and qty > 0 and key not in seen:
                seen.add(key)
                devices.append({"name": desc, "quantity": qty, "section": ""})

        return devices

    def _parse_fbm_format(self, df: pd.DataFrame) -> list[dict]:
        """
        Parse the specific FbM Tech Estimation layout:
        Col 0: ISA level  | Col 1: Section  | Col 2: CAPEX/OPEX code
        Col 3: Description | … | Quantity column (identified by scanning)
        """
        devices: list[dict] = []
        qty_col = self._find_qty_column(df)
        if qty_col is None:
            return devices

        current_section = ""
        seen: set[str] = set()
        for row_idx in range(len(df)):
            row = df.iloc[row_idx].fillna("")
            str_vals = [str(v).strip() for v in row]

            # Description is in index 3 (FbM standard layout)
            desc = str_vals[3] if len(str_vals) > 3 else ""
            qty_raw = str_vals[qty_col] if qty_col < len(str_vals) else ""

            # Track section headers (col 1 usually contains section names)
            if len(str_vals) > 1 and str_vals[1] and not self._is_numeric(str_vals[1]):
                candidate = str_vals[1]
                if len(candidate) > 3 and not candidate.upper().startswith(("CAPEX", "OPEX")):
                    current_section = candidate

            key = desc.lower()
            if self._is_valid_device(desc, qty_raw) and key not in seen:
                seen.add(key)
                devices.append({
                    "name": desc,
                    "quantity": int(float(qty_raw)),
                    "section": current_section,
                })

        return devices

    def _find_qty_column(self, df: pd.DataFrame) -> int | None:
        """
        Identify the Quantity column index in a FbM estimation sheet.
        The Qty column is the first column whose header contains 'qty' or 'quantity',
        or – if no header row – the column with the most small positive integers.
        """
        # Try header-based detection in first 20 rows
        for row_idx in range(min(20, len(df))):
            vals = [str(v).strip().lower() for v in df.iloc[row_idx].fillna("")]
            for i, v in enumerate(vals):
                if "qty" in v or "quantity" in v:
                    return i

        # Fallback: column with most positive integers
        best_col = None
        best_count = 0
        for col in df.columns:
            count = 0
            for val in df[col].fillna(""):
                s = str(val).strip()
                if self._is_numeric(s) and 0 < int(float(s)) < 1000:
                    count += 1
            if count > best_count:
                best_count = count
                best_col = col

        return best_col  # May still be None if no suitable column found

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _is_numeric(val: str) -> bool:
        try:
            f = float(val)
            return f == int(f) and f >= 0
        except (ValueError, TypeError):
            return False

    @staticmethod
    def _looks_like_device(text: str) -> bool:
        """Rough heuristic: device names are mixed-case, not all-caps."""
        return not text.isupper() and len(text) > 3

    @staticmethod
    def _is_valid_device(desc: Any, qty_raw: Any) -> bool:
        desc_str = str(desc).strip()
        if not desc_str or desc_str.lower() in ("nan", "", "description", "item", "n/a"):
            return False
        if desc_str.upper().startswith(("CAPEX", "OPEX")):
            return False
        try:
            qty = int(float(qty_raw))
        except (ValueError, TypeError):
            return False
        return qty > 0 and len(desc_str) > 3
