"""
parser_engine.py

Parser engine for Universal ATP Generator (FTTH).
Implements parsing contracts specified in MASTER PROMPT (§7, §15):
- parse_fat_inline(commands: list[str]) -> list[str]
- parse_pole_inline(commands: list[str]) -> list[dict]
- parse_txt_upload(content: str, upload_type: str) -> list[dict]
- normalize_json_filename(nama_lokasi: str) -> str

The functions aim to be defensive and provide clear line-level error messages for .txt uploads.
"""

from __future__ import annotations
import re
from typing import List, Dict, Any, Optional

from system_config import json_safe_filename


# -----------------------------
# parse_fat_inline
# -----------------------------
RE_FAT_CMD = re.compile(r"^([A-Za-z]+)(\d+)$")


def parse_fat_inline(commands: List[str]) -> List[str]:
    """
    Parse inline FAT commands into a sequential list of FAT names (zero-padded 2 digits).

    Rules (per spec §7.1):
    - Regex: ^([A-Za-z]+)(\d+)$ -> group1 letters (line), group2 digits (count)
    - Output: list of strings like 'A01','A02', ...
    - Supports multiple commands; results are concatenated in order.
    - Fallback: if a command does not match the regex, return the trimmed command "as-is".

    Example:
        ['A12'] -> ['A01','A02',...,'A12']
        ['a3','B2'] -> ['A01','A02','A03','B01','B02']
    """
    result: List[str] = []
    for raw in commands:
        if not raw:
            continue
        line = raw.strip()
        # support commands that contain whitespace/newline separated values
        # split by whitespace and commas if user pasted multiple on one line
        candidates = re.split(r"[\n,;]+", line)
        for cand in candidates:
            cand = cand.strip()
            if not cand:
                continue
            m = RE_FAT_CMD.match(cand)
            if m:
                letters = m.group(1).upper()
                try:
                    count = int(m.group(2))
                except ValueError:
                    # fallback to raw
                    result.append(cand)
                    continue
                for i in range(1, count + 1):
                    # Zero-pad 2 digits, return only the code (e.g., A01)
                    result.append(f"{letters}{i:02d}")
            else:
                # fallback: include as-is (trimmed)
                result.append(cand)
    return result


# -----------------------------
# parse_pole_inline
# -----------------------------
RE_POLE_SIZE = re.compile(r"^(\d)([\d\.]*)$")


def _convert_pole_size(raw: str) -> Optional[str]:
    """
    Convert numeric pole spec like '73' or '72.5' into '7 METER 3 INCH' or '7 METER 2.5 INCH'.
    If unable to parse, return None.
    Behavior follows §7.2.
    """
    raw = raw.strip()
    m = RE_POLE_SIZE.match(raw)
    if not m:
        return None
    meter = m.group(1)
    inch = m.group(2)
    if inch == "":
        inch = "0"
    return f"{int(meter)} METER {inch} INCH"


def parse_pole_inline(commands: List[str]) -> List[Dict[str, Any]]:
    """
    Parse pole commands into structured dicts.

    Input examples:
      - 'pole 73=3'
      - 'ext 74=2'
      - Multi-line supported. Leading/trailing whitespace trimmed.

    Output per item (dict): {
        'title': 'Pole Erection 73' or 'Pole Erection EXT 74',
        'type': 'NEW POLE' | 'EXT POLE',
        'size_clean': '7 METER 3 INCH',
        'qty': 3
    }

    Rules derived from spec §7.2.
    """
    result: List[Dict[str, Any]] = []
    for raw in commands:
        if not raw:
            continue
        # allow multiple commands in a single string separated by newlines or commas
        lines = re.split(r"[\n,;]+", raw)
        for line in lines:
            line = line.strip()
            if not line:
                continue
            # Expect format: '<keyword> <value>=<qty>' or '<keyword><space><value>=<qty>'
            if '=' not in line:
                # fallback: try to parse simple 'pole 73' (qty 1)
                parts = line.split()
                if len(parts) >= 2 and parts[0].lower() in ("pole", "ext"):
                    keyword = parts[0].lower()
                    raw_value = parts[1]
                    qty = 1
                else:
                    # can't parse -> skip
                    continue
            else:
                left, right = line.split('=', 1)
                qty_raw = right.strip()
                try:
                    qty = int(qty_raw)
                except Exception:
                    # invalid qty -> skip this line
                    continue
                left = left.strip()
                left_parts = left.split()
                if len(left_parts) == 0:
                    continue
                keyword = left_parts[0].lower()
                raw_value = left_parts[1] if len(left_parts) > 1 else ''

            if keyword == 'pole':
                pole_type = 'NEW POLE'
                title = f"Pole Erection {raw_value}"
            elif keyword == 'ext':
                pole_type = 'EXT POLE'
                title = f"Pole Erection EXT {raw_value}"
            else:
                # unknown keyword: treat as NEW POLE but preserve raw
                pole_type = 'NEW POLE'
                title = f"Pole Erection {raw_value or keyword}"

            size_converted = _convert_pole_size(raw_value) or raw_value
            result.append({
                'title': title,
                'type': pole_type,
                'size_clean': size_converted,
                'qty': qty,
            })
    return result


# -----------------------------
# parse_txt_upload
# -----------------------------

# Supported upload_type values (as keys):
SUPPORTED_UPLOAD_TYPES = {
    'otdr_cluster': 'OTDR_CLUSTER',
    'otdr_subfeeder': 'OTDR_SUBFEEDER',
    'opm_feeder': 'OPM_FEEDER',
    'opm_distribution': 'OPM_DISTRIBUTION',
}


def _error(line_no: int, msg: str) -> ValueError:
    return ValueError(f"Line {line_no}: {msg}")


def parse_txt_upload(content: str, upload_type: str) -> List[Dict[str, Any]]:
    """
    Parse uploaded .txt content according to upload_type.

    Parameters:
      - content: raw string of the uploaded file
      - upload_type: one of 'otdr_cluster', 'otdr_subfeeder', 'opm_feeder', 'opm_distribution'

    Returns: list of dict rows parsed into canonical keys suitable for merging into session_state.

    Raises ValueError with detailed per-line diagnostics if validation fails.

    Behavior summarized from §9.5
    """
    if upload_type not in SUPPORTED_UPLOAD_TYPES:
        raise ValueError(f"Unsupported upload_type '{upload_type}'. Supported: {list(SUPPORTED_UPLOAD_TYPES.keys())}")

    rows: List[Dict[str, Any]] = []
    lines = content.splitlines()

    for idx, raw in enumerate(lines, start=1):
        line = raw.strip()
        if not line:
            continue
        if line.startswith('#'):
            continue
        # split by '|' and strip each field
        parts = [p.strip() for p in line.split('|')]
        # convert empty strings to '' explicitly
        parts = [p if p != '' else '' for p in parts]

        t = upload_type
        try:
            if t == 'otdr_cluster':
                # Expect 5 columns: FAT_NAME | DISTANCE_1310 | LOSS_1310 | DISTANCE_1550 | LOSS_1550
                if len(parts) != 5:
                    raise _error(idx, f"Expected 5 columns for OTDR CLUSTER, got {len(parts)}")
                fat_name, d1310, l1310, d1550, l1550 = parts
                rows.append({
                    'FAT_NAME': fat_name,
                    'DISTANCE_1310': d1310,
                    'LOSS_1310': l1310,
                    'DISTANCE_1550': d1550,
                    'LOSS_1550': l1550,
                })

            elif t == 'otdr_subfeeder':
                # Expect 5 columns: CORE_NO | DISTANCE_1310 | LOSS_1310 | DISTANCE_1550 | LOSS_1550
                if len(parts) != 5:
                    raise _error(idx, f"Expected 5 columns for OTDR SUBFEEDER, got {len(parts)}")
                core_no, d1310, l1310, d1550, l1550 = parts
                rows.append({
                    'CORE_NO': core_no,
                    'DISTANCE_1310': d1310,
                    'LOSS_1310': l1310,
                    'DISTANCE_1550': d1550,
                    'LOSS_1550': l1550,
                })

            elif t == 'opm_feeder':
                # Expect at least 3 columns: SPL_ID | SPL_SN | OPM_BEFORE | P1 | P2 | ... up to P16
                if len(parts) < 3:
                    raise _error(idx, f"Expected at least 3 columns for OPM FEEDER, got {len(parts)}")
                spl_id = parts[0]
                spl_sn = parts[1]
                opm_before = parts[2]
                port_values = parts[3:]
                # map ports to P1..Pn
                row: Dict[str, Any] = {
                    'SPL_ID': spl_id,
                    'SPL_SN': spl_sn,
                    'OPM_BEFORE': opm_before,
                }
                for i, val in enumerate(port_values, start=1):
                    key = f'P{i}'
                    row[key] = val
                rows.append(row)

            elif t == 'opm_distribution':
                # Expect at least 2 columns: FAT_NAME | P1 | P2 | ...
                if len(parts) < 2:
                    raise _error(idx, f"Expected at least 2 columns for OPM DISTRIBUTION, got {len(parts)}")
                fat_name = parts[0]
                port_values = parts[1:]
                row = {'FAT_NAME': fat_name}
                for i, val in enumerate(port_values, start=1):
                    key = f'P{i}'
                    row[key] = val
                rows.append(row)

            else:
                raise _error(idx, f"Unhandled upload_type '{t}'")
        except ValueError as ve:
            # re-raise with full context
            raise ve
        except Exception as exc:
            raise _error(idx, f"Unexpected parse error: {exc}")

    return rows


# -----------------------------
# normalize_json_filename
# -----------------------------

def normalize_json_filename(nama_lokasi: str) -> str:
    """
    Normalize nama_lokasi into a JSON filename per spec §12/§13.
    Uses system_config.json_safe_filename() and appends .json if missing.

    Examples:
      'DUSUN BOGO RW 08 FDT-2' -> 'DUSUN_BOGO_RW_08_FDT-2.json'
    """
    base = json_safe_filename(nama_lokasi)
    if not base.lower().endswith('.json'):
        return f"{base}.json"
    return base


# -----------------------------
# Module quick-test helpers
# -----------------------------
if __name__ == '__main__':
    # quick interactive smoke tests
    print(parse_fat_inline(['A3', 'B2']))
    print(parse_pole_inline(['pole 73=3', 'ext 74=2']))
    sample_otdr = """
# FAT_NAME | DISTANCE_1310 | LOSS_1310 | DISTANCE_1550 | LOSS_1550
FAT A01 | 0.194 | 1.000 | 0.196 | 1.340
FAT A02 | 0.201 | 0.980 | 0.203 | 1.310
"""
    print(parse_txt_upload(sample_otdr, 'otdr_cluster'))
    print(normalize_json_filename('DUSUN BOGO RW 08 FDT-2'))
