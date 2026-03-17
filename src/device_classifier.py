"""
Device classifier – categorises extracted devices into:
  IT | OT | Network | MDF | Automation | Software | Other

Uses keyword matching (no API call required, instant classification).
"""
from __future__ import annotations

# ── Keyword lookup tables (all lowercase) ────────────────────────────────────

_MDF_KW = {
    "core switch",
    "distribution switch",
    "mdf rack",
    "mdf w/patch",
    "rack w/patch",
    "server rack",
    "24ru rack",
    "42ru rack",
    "37ru rack",
    "main distribution frame",
    "2200va",
    "ups w/warranty (2200",
    "ups.*2200",
}

_AUTOMATION_KW = {
    "cubiscan",
    "rfid tunnel",
    "rfid portal",
    "panda",
    "print & apply",
    "print apply",
    "automated print",
    "carton erect",
    "carton closing",
    "sealing machine",
    "conveyor",
    "ot machine",
    "automation machine",
    "scanner station",
    "weighing",
    "dimensioning",
}

_OT_KW = {
    "rf scanner",
    "radio frequency scanner",
    "standard radio frequency",
    "bluetooth ring scanner",
    "bluetooth ring charger",
    "rfid scanner",
    "rfid hand",
    "ultra-rugged barcode",
    "smartphone barcode",
    "vehicle mounted computer",
    "cold storage",
    "wired barcode scanner",
    "hand scanner",
    "handheld scanner",
    "barcode scanner",
    "industrial label printer",
    "label printer",
    "mobile printer",
    "mobile label",
    "zebra",
    "honeywell ct",
    "honeywell ck",
    "honeywell vm",
    "honeywell rt",
    "honeywell eda",
    "rugged tablet",
    "forklift tablet",
    "tablet (yard",
    "tablet devices",
    "battery charger",
    "quad battery charger",
    "extra battery",
    "soti",
    "mdm",
}

_NETWORK_KW = {
    "access switch",
    "ethernet switch",
    "managed switch",
    "cisco catalyst",
    "access point",
    "wireless ap",
    "wlan",
    "patch panel",
    "structured cabling",
    "cable installation",
    "network cabling",
    "vap mounting",
    "network hardware installation",
    "idf cabinet",
    "idf cabinets",
    "firewall",
    "security appliance",
    "router",
    "sfp",
    "transceiver",
    "wan link",
    "maersknet",
    "raw business",
    "virtual hardware",
    "vrb",
    "1400va",
    "ups w/warranty (1400",
}

_IT_KW = {
    "laptop computer",
    "laptop",
    "notebook",
    "elitebook",
    "latitude",
    "thinkpad",
    "surface",
    "desktop computer",
    "desktop pc",
    "elitedesk",
    "optiplex",
    "mini pc",
    "monitor",
    "display",
    "docking station",
    "dock station",
    "keyboard",
    "mouse",
    "headset",
    "headphone",
    "laser printer",
    "mfp",
    "multifunction",
    "all in one printer",
    "mono printer",
    "cellphone",
    "cell phone",
    "mobile phone",
    "smartphone",
    "ups w/warranty",
    "meeting room",
    "video bar",
    "conferencing pc",
    "polycom",
    "jabra",
    "poly",
    "pancast",
    "webcam",
    "camera",
    "projector",
    "microsoft teams",
    "o365",
    "server room",
    "aot server",
}

_SOFTWARE_KW = {
    "license",
    "licence",
    "subscription",
    "software",
    "hosting",
    "configuration cost",
    "integration cost",
    "support cost",
    "annual support",
    "bartender",
    "wms",
    "dms",
    "nci",
    "nkp",
}


class DeviceClassifier:
    """Classify devices into hardware/software categories via keyword matching."""

    CATEGORIES = ["IT", "OT", "Network", "MDF", "Automation", "Software", "Other"]

    def classify(self, device_name: str) -> str:
        """Return the category for a single device name string."""
        n = device_name.lower()

        # Check in priority order: most specific first
        for kw in _MDF_KW:
            if kw in n:
                return "MDF"
        for kw in _AUTOMATION_KW:
            if kw in n:
                return "Automation"
        for kw in _OT_KW:
            if kw in n:
                return "OT"
        for kw in _NETWORK_KW:
            if kw in n:
                return "Network"
        for kw in _IT_KW:
            if kw in n:
                return "IT"
        for kw in _SOFTWARE_KW:
            if kw in n:
                return "Software"

        return "Other"

    def classify_all(
        self, devices: list[dict]
    ) -> dict[str, list[dict]]:
        """
        Classify all devices and return a dict grouped by category.

        Each device dict gets an added 'category' key.
        """
        result: dict[str, list[dict]] = {cat: [] for cat in self.CATEGORIES}
        for device in devices:
            cat = self.classify(device["name"])
            result[cat].append({**device, "category": cat})
        return result
