"""
system_config.py

Rulebook & konstanta untuk Universal ATP Generator (FTTH).
Isi ini mengikuti MASTER PROMPT — semua nilai dan helper penting dikumpulkan di satu tempat.

CATATAN:
- Jangan ubah nama konstanta utama tanpa memperbarui modul lain.
- PORT_COUNT default = 8, mendukung 8 atau 16 (konfigurable).
- RED_TAB_RGB sesuai contoh: 'FFFF0000' (ARGB hex string openpyxl style).
"""

from __future__ import annotations
import re
from typing import Iterable, List, Dict

# ---------- Port / IOR Defaults ----------
DEFAULT_PORT_COUNT: int = 8
SUPPORTED_PORT_COUNTS = (8, 16)

DEFAULT_IOR_1310: float = 1.4681
DEFAULT_IOR_1550: float = 1.4676

# ---------- Red Tab Safeguard (openpyxl ARGB string) ----------
# Example in spec: tabColor.rgb == "FFFF0000"
RED_TAB_RGB = "FFFF0000"

# ---------- Sheet name patterns / ranges (template naming) ----------
# These regexes used to detect template sheet groups that will be processed/renamed/spilled.
RE_FAT_SHEETS = re.compile(r"^FAT_[0-9]{3}$", re.IGNORECASE)
RE_POLE_SHEETS = re.compile(r"^POLE_[0-9]{3}$", re.IGNORECASE)
RE_BA_SPLITTER_FAT = re.compile(r"^BA Splitter FAT_[0-9]{3}$", re.IGNORECASE)
RE_OTDR_SUMMARY_FDT_FAT = re.compile(r"^OTDR Sumary \(FDT-FAT\)_[0-9]{3}$", re.IGNORECASE)  # note: 'Sumary' typo preserved
RE_OPM_DISTRIBUTION = re.compile(r"^OPM_DISTRIBUTION_[0-9]{3}$", re.IGNORECASE)
RE_OPM_SHEETS = re.compile(r"^OPM_[0-9]{3}$", re.IGNORECASE)
RE_OTDR_WAVE = re.compile(r"^OTDR Summary \(WAVE\)_[0-9]{3}$", re.IGNORECASE)

# Mapping kapasitas kaku per group (lihat §6.1)
SHEET_GROUP_CAPACITY: Dict[str, int] = {
    "FAT": 50,  # FAT_001 .. FAT_050
    "POLE": 20,  # POLE_001 .. POLE_020
    "BA_SPLITTER_FAT": 20,  # 20 sheet, 10 baris FAT per sheet
    "OTDR_SUMMARY_FDT_FAT": 40,  # 40 slot FAT per sheet (20 kiri + 20 kanan)
    "OPM_DISTRIBUTION": 10,  # 10 slot FAT per sheet (5 kol x 2 baris)
    "OPM": 4,  # 4 blok split per sheet (per sheet capacity)
    "OTDR_WAVE": 9999,  # handled specially (split 1310/1550)
}

# ---------- Token lists (Phase 1 & Phase 2) ----------
# Phase-1 tokens (administrasi, FAT/pole placeholders)
PHASE1_TOKENS = (
    "[NAMA_PROYEK]",
    "[REGION]",
    "[NAMA_LOKASI]",
    "[ID_LOKASI]",
    "[ALAMAT]",
    "[NAMA_OLT]",
    "[ID_FDT_FROM]",
    "[ID_FAT_TO]",
    "[NAMA_PT_VENDOR]",
    "[REP_VENDOR]",
    "[JABATAN_VENDOR]",
    "[NAMA_PT_CUSTOMER]",
    "[REP_CUSTOMER]",
    "[JABATAN_CUSTOMER]",
    "[TANGGAL_TEST]",
    "[NO_PO]",
    # Fase1 dynamic placeholders:
    "[INPUT_FAT_NAME]",
    "[INPUT_POLE_DESC]",
    "[INPUT_POLE_SIZE]",
)

# Phase-2 tokens (technical measurement placeholders)
PHASE2_TOKENS = (
    "[SPL_ID]",
    "[SPL_SN]",
    "[OPM_BEFORE]",
    "[OPM_AFTER]",
    "[INPUT_OPM]",
    "[DISTANCE]",
    "[1310]",
    "[1550]",
    "[WAVE_LENGHT]",
    "[DISTANCE_SF]",
    "[OTDR_SF]",
    "[IOR_1310]",
    "[IOR_1550]",
)

# Combined token set for convenience
ALL_TOKENS = tuple(set(PHASE1_TOKENS) | set(PHASE2_TOKENS))

# ---------- Blacklist per mode ----------
# Sheets/operations that must be excluded or handled differently per mode.
BLACKLIST = {
    "cluster": {
        # sheets that are irrelevant for cluster mode (and should be removed)
        "remove_patterns": [
            re.compile(r"^OTDR Summary 1310$", re.IGNORECASE),
            re.compile(r"^OTDR Summary 1550$", re.IGNORECASE),
            re.compile(r"^E2E OPM Distribution", re.IGNORECASE),  # if spec differs in nomenclature
        ],
        # sheets that must be preserved but not filled in cluster mode (if any)
        "preserve_patterns": [],
    },
    "subfeeder": {
        # sheets that are irrelevant for subfeeder mode (and should be removed)
        "remove_patterns": [
            re.compile(r"^BA Splitter FAT", re.IGNORECASE),
            re.compile(r"^E2E OPM Distribution", re.IGNORECASE),
            re.compile(r"^FAT_[0-9]{3}$", re.IGNORECASE),  # multiple FAT sheets in cluster mode
        ],
        "preserve_patterns": [],
    },
}

# ---------- Cover sheet detection ----------
COVER_KEYWORD = "cover"


def is_cover_sheet(sheet_name: str) -> bool:
    """
    Return True if sheet name is a Cover sheet (case-insensitive containment).
    Matches any sheet whose name contains 'cover' (e.g. 'ATP CW Cover', 'Cover CW OPM').
    """
    if not sheet_name:
        return False
    return COVER_KEYWORD in sheet_name.lower()


def sheet_has_red_tab(ws) -> bool:
    """
    Check worksheet tab color for red tab safeguard.
    Expects an openpyxl worksheet object; returns True if tabColor.rgb matches RED_TAB_RGB.
    Caller should handle attribute absence safely.
    """
    try:
        tab_color = ws.sheet_properties.tabColor
        if tab_color is None:
            return False
        # openpyxl may store rgb as 'FFFF0000' or similar; compare uppercase
        rgb = getattr(tab_color, "rgb", None)
        if rgb is None:
            # Some templates use themeColor or indexed color — treat as not-red (safe)
            return False
        return str(rgb).upper() == RED_TAB_RGB.upper()
    except Exception:
        # Fail-safe: if any unexpected structure, do not claim red (so higher-level code can decide)
        return False


# ---------- Helper utilities ----------
def validate_port_count(port_count: int) -> int:
    """
    Validate and normalize port_count. Return supported value or raise ValueError.
    """
    if port_count in SUPPORTED_PORT_COUNTS:
        return port_count
    raise ValueError(f"Unsupported PORT_COUNT {port_count}. Supported: {SUPPORTED_PORT_COUNTS}")


def normalize_mode(mode: str) -> str:
    """
    Normalize mode slug to canonical 'cluster' or 'subfeeder'.
    """
    if not mode:
        raise ValueError("mode must be provided")
    m = mode.strip().lower()
    if m.startswith("cl"):
        return "cluster"
    if m.startswith("s"):
        return "subfeeder"
    raise ValueError("mode must be 'cluster' or 'subfeeder'")


# ---------- Sheet naming helpers (nomenclature) ----------
def fmt_fat_name(line_letter: str, index: int) -> str:
    """
    Zero-pad 2-digit FAT name e.g. ('A', 1) -> 'FAT A01'
    Enforce the Zero-Pad Wajib 2 Digit rule (no A010).
    """
    return f"FAT {line_letter.upper()}{index:02d}"


def json_safe_filename(nama_lokasi: str) -> str:
    """
    Convert nama_lokasi to JSON-safe filename per spec:
    - spaces -> underscore
    - special chars (except dash and underscore) -> underscore
    - preserve leading zeros and dashes as in prompt (but replace '.' too)
    """
    if not nama_lokasi:
        return "project.json"
    # preserve '-' per example (they kept '-') but prompt earlier suggested replace '-' and '.' -> '_'
    # Spec: replace special characters -> '_' (example shows '-' kept in file sample but §13 says special '-' -> '_')
    # Follow §12: replace characters special with '_', keep leading zero and spaces replaced by '_'
    s = nama_lokasi.strip()
    # replace spaces with underscore
    s = s.replace(" ", "_")
    # replace any character not alnum, underscore, or dash with underscore
    s = re.sub(r"[^A-Za-z0-9_\-]", "_", s)
    return s


# ---------- Export list ----------
__all__ = [
    "DEFAULT_PORT_COUNT",
    "SUPPORTED_PORT_COUNTS",
    "DEFAULT_IOR_1310",
    "DEFAULT_IOR_1550",
    "RED_TAB_RGB",
    "RE_FAT_SHEETS",
    "RE_POLE_SHEETS",
    "RE_BA_SPLITTER_FAT",
    "RE_OTDR_SUMMARY_FDT_FAT",
    "RE_OPM_DISTRIBUTION",
    "RE_OPM_SHEETS",
    "RE_OTDR_WAVE",
    "SHEET_GROUP_CAPACITY",
    "PHASE1_TOKENS",
    "PHASE2_TOKENS",
    "ALL_TOKENS",
    "BLACKLIST",
    "is_cover_sheet",
    "sheet_has_red_tab",
    "validate_port_count",
    "normalize_mode",
    "fmt_fat_name",
    "json_safe_filename",
]
