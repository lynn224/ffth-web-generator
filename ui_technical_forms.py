"""
ui_technical_forms.py

Streamlit render helpers for Phase 2 forms and autosave integration.
Implements required render functions and autosave/load_autosave per contract (§9, §15):
- render_ior_inputs()
- render_modul_2a_splitter()
- render_modul_2b_opm_dist(parsed_fat: list)
- render_modul_2c_2d_otdr(parsed_fat: list, mode: str)
- auto_save(username: str, project_slug: str, state: dict)
- load_autosave(username: str, project_slug: str) -> dict | None

This module avoids any heavy I/O on import and uses Streamlit APIs at render time.
"""

from __future__ import annotations
import os
import json
import time
from typing import List, Dict, Any, Optional

import streamlit as st
import pandas as pd

from parser_engine import parse_txt_upload
from system_config import DEFAULT_IOR_1310, DEFAULT_IOR_1550, validate_port_count

HISTORY_DIR = "history_database"
AUTOSAVE_DIR = os.path.join(HISTORY_DIR, "autosave")
os.makedirs(AUTOSAVE_DIR, exist_ok=True)


# -----------------------------
# Utility helpers
# -----------------------------

def _autosave_path(username: str, project_slug: str) -> str:
    safe_user = username or "anonymous"
    safe_slug = project_slug or "untitled"
    filename = f"autosave_{safe_user}_{safe_slug}.json"
    return os.path.join(AUTOSAVE_DIR, filename)


# -----------------------------
# IOR Inputs
# -----------------------------

def render_ior_inputs():
    """Render two numeric inputs for IOR and store values in session_state.
    Keys used: st.session_state['IOR_1310'], st.session_state['IOR_1550']
    """
    col1, col2 = st.columns(2)
    with col1:
        ior1310 = st.number_input("IOR 1310 nm", value=st.session_state.get("IOR_1310", DEFAULT_IOR_1310), format="%.6f", key="IOR_1310")
    with col2:
        ior1550 = st.number_input("IOR 1550 nm", value=st.session_state.get("IOR_1550", DEFAULT_IOR_1550), format="%.6f", key="IOR_1550")
    st.session_state["IOR_1310"] = float(ior1310)
    st.session_state["IOR_1550"] = float(ior1550)


# -----------------------------
# Splitter module (2A)
# -----------------------------

OPM_FEEDER_TEMPLATE_TXT = """# FORMAT UPLOAD OPM FEEDER
# SPL_ID | SPL_SN | OPM_BEFORE | P1 | P2 | P3 | P4 | P5 | P6 | P7 | P8
# Example:
SPLITTER 1 | SN-BMI-001 | -12.15 | -18.10 | -18.25 | | | | | |
"""


def _make_splitter_dataframe(existing: Optional[List[Dict[str, Any]]] = None, port_count: int = 8) -> pd.DataFrame:
    cols = ["SPL_ID", "SPL_SN", "OPM_BEFORE"] + [f"P{i}" for i in range(1, port_count+1)]
    if existing:
        df = pd.DataFrame(existing)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
    else:
        df = pd.DataFrame(columns=cols)
    return df


def render_modul_2a_splitter(key_prefix: str = "splitter", port_count: int = 8):
    """Render OPM Feeder / Splitter grid with upload/download .txt support.

    Stores dataframe in session_state['df_splitter_fdt'] by default.
    """
    validate_port_count(port_count)
    st.header("OPM Feeder / Splitter")
    col1, col2 = st.columns([3, 1])
    with col1:
        data_key = "df_splitter_fdt"
        if data_key not in st.session_state:
            st.session_state[data_key] = _make_splitter_dataframe(None, port_count).to_dict(orient="records")
        df = pd.DataFrame(st.session_state.get(data_key, []))
        edited = st.data_editor(df, num_rows="dynamic", key=f"editor_{data_key}")
        st.session_state[data_key] = edited.to_dict(orient="records")
    with col2:
        st.download_button("Download Template .txt", data=OPM_FEEDER_TEMPLATE_TXT, file_name="template_opm_feeder.txt", mime="text/plain")
        uploaded = st.file_uploader("Upload dari File .txt", type=["txt"], key=f"upload_{key_prefix}")
        if uploaded is not None:
            content = uploaded.read().decode("utf-8")
            try:
                rows = parse_txt_upload(content, 'opm_feeder')
            except Exception as e:
                st.error(f"Error parsing file: {e}")
                rows = []
            if rows:
                # merge rows: existing rows are preserved unless replaced
                existing = {r.get('SPL_ID', f'__{i}'): r for i, r in enumerate(st.session_state.get(data_key, []))}
                for r in rows:
                    key = r.get('SPL_ID') or f"__new_{len(existing)}"
                    existing[key] = r
                st.session_state[data_key] = list(existing.values())
                st.success(f"Parsed {len(rows)} rows and merged into grid.")


# -----------------------------
# OPM Distribution module (2B)
# -----------------------------

OPM_DISTR_TEMPLATE_TXT = """# FORMAT UPLOAD OPM DISTRIBUSI
# FAT_NAME | P1 | P2 | P3 | P4 | P5 | P6 | P7 | P8
# Example:
FAT A01 | -19.10 | -19.15 | | | | | |
"""


def _make_opm_dist_dataframe(existing: Optional[List[Dict[str, Any]]], parsed_fat: List[str], port_count: int = 8) -> pd.DataFrame:
    cols = ["FAT_NAME"] + [f"P{i}" for i in range(1, port_count+1)]
    if existing:
        df = pd.DataFrame(existing)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
    else:
        # Create fixed rows from parsed_fat
        rows = [{"FAT_NAME": name, **{f"P{i}": "" for i in range(1, port_count+1)}} for name in parsed_fat]
        df = pd.DataFrame(rows, columns=cols)
    return df


def render_modul_2b_opm_dist(parsed_fat: List[str], port_count: int = 8):
    validate_port_count(port_count)
    st.header("OPM Distribution")
    data_key = "df_opm_distribution"
    if data_key not in st.session_state:
        st.session_state[data_key] = _make_opm_dist_dataframe(None, parsed_fat, port_count).to_dict(orient="records")
    df = pd.DataFrame(st.session_state.get(data_key, []))
    edited = st.data_editor(df, num_rows="fixed", key=f"editor_{data_key}")
    st.session_state[data_key] = edited.to_dict(orient="records")

    col1, col2 = st.columns([3, 1])
    with col2:
        st.download_button("Download Template .txt", data=OPM_DISTR_TEMPLATE_TXT, file_name="template_opm_distribution.txt", mime="text/plain")
        uploaded = st.file_uploader("Upload dari File .txt", type=["txt"], key="upload_opm_dist")
        if uploaded is not None:
            content = uploaded.read().decode("utf-8")
            try:
                rows = parse_txt_upload(content, 'opm_distribution')
            except Exception as e:
                st.error(f"Error parsing file: {e}")
                rows = []
            if rows:
                # merge by FAT_NAME
                existing = {r.get('FAT_NAME'): r for r in st.session_state.get(data_key, [])}
                for r in rows:
                    name = r.get('FAT_NAME')
                    if name:
                        existing[name] = {**existing.get(name, {}), **r}
                st.session_state[data_key] = list(existing.values())
                st.success(f"Parsed {len(rows)} rows and merged into grid.")


# -----------------------------
# OTDR modules (2C/2D)
# -----------------------------

OTDR_CLUSTER_TEMPLATE_TXT = """# FORMAT UPLOAD OTDR CLUSTER
# FAT_NAME | DISTANCE_1310 | LOSS_1310 | DISTANCE_1550 | LOSS_1550
# Example:
FAT A01 | 0.194 | 1.000 | 0.196 | 1.340
"""

OTDR_SUBFEEDER_TEMPLATE_TXT = """# FORMAT UPLOAD OTDR SUBFEEDER
# CORE_NO | DISTANCE_1310 | LOSS_1310 | DISTANCE_1550 | LOSS_1550
# Example:
Core 1 | 0.520 | 0.040 | 0.530 | 0.050
"""


def _make_otdr_cluster_df(existing: Optional[List[Dict[str, Any]]], parsed_fat: List[str]) -> pd.DataFrame:
    cols = ["FAT_NAME", "DISTANCE_1310", "LOSS_1310", "DISTANCE_1550", "LOSS_1550"]
    if existing:
        df = pd.DataFrame(existing)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
    else:
        rows = [{"FAT_NAME": name, "DISTANCE_1310": "", "LOSS_1310": "", "DISTANCE_1550": "", "LOSS_1550": ""} for name in parsed_fat]
        df = pd.DataFrame(rows, columns=cols)
    return df


def _make_otdr_subfeeder_df(existing: Optional[List[Dict[str, Any]]]) -> pd.DataFrame:
    cols = ["CORE_NO", "DISTANCE_1310", "LOSS_1310", "DISTANCE_1550", "LOSS_1550"]
    if existing:
        df = pd.DataFrame(existing)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
    else:
        df = pd.DataFrame(columns=cols)
    return df


def render_modul_2c_2d_otdr(parsed_fat: List[str], mode: str = 'cluster'):
    st.header("OTDR Measurements")
    mode = mode.lower()

    if mode == 'cluster':
        data_key = 'df_otdr_cluster'
        if data_key not in st.session_state:
            st.session_state[data_key] = _make_otdr_cluster_df(None, parsed_fat).to_dict(orient='records')
        df = pd.DataFrame(st.session_state.get(data_key, []))
        edited = st.data_editor(df, num_rows='fixed', key=f'editor_{data_key}')
        st.session_state[data_key] = edited.to_dict(orient='records')

        col1, col2 = st.columns([3,1])
        with col2:
            st.download_button("Download Template .txt", data=OTDR_CLUSTER_TEMPLATE_TXT, file_name='template_otdr_cluster.txt', mime='text/plain')
            uploaded = st.file_uploader("Upload dari File .txt (Cluster)", type=['txt'], key='upload_otdr_cluster')
            if uploaded is not None:
                content = uploaded.read().decode('utf-8')
                try:
                    rows = parse_txt_upload(content, 'otdr_cluster')
                except Exception as e:
                    st.error(f"Error parsing file: {e}")
                    rows = []
                if rows:
                    # merge by FAT_NAME
                    existing = {r.get('FAT_NAME'): r for r in st.session_state.get(data_key, [])}
                    for r in rows:
                        name = r.get('FAT_NAME')
                        if name:
                            existing[name] = {**existing.get(name, {}), **r}
                    st.session_state[data_key] = list(existing.values())
                    st.success(f"Parsed {len(rows)} rows and merged into OTDR cluster grid.")

    else:
        data_key = 'df_otdr_subfeeder'
        if data_key not in st.session_state:
            st.session_state[data_key] = _make_otdr_subfeeder_df(None).to_dict(orient='records')
        df = pd.DataFrame(st.session_state.get(data_key, []))
        edited = st.data_editor(df, num_rows='dynamic', key=f'editor_{data_key}')
        st.session_state[data_key] = edited.to_dict(orient='records')

        col1, col2 = st.columns([3,1])
        with col2:
            st.download_button("Download Template .txt", data=OTDR_SUBFEEDER_TEMPLATE_TXT, file_name='template_otdr_subfeeder.txt', mime='text/plain')
            uploaded = st.file_uploader("Upload dari File .txt (Subfeeder)", type=['txt'], key='upload_otdr_subfeeder')
            if uploaded is not None:
                content = uploaded.read().decode('utf-8')
                try:
                    rows = parse_txt_upload(content, 'otdr_subfeeder')
                except Exception as e:
                    st.error(f"Error parsing file: {e}")
                    rows = []
                if rows:
                    # merge by CORE_NO
                    existing = {r.get('CORE_NO'): r for r in st.session_state.get(data_key, [])}
                    for r in rows:
                        name = r.get('CORE_NO')
                        if name:
                            existing[name] = {**existing.get(name, {}), **r}
                    st.session_state[data_key] = list(existing.values())
                    st.success(f"Parsed {len(rows)} rows and merged into OTDR subfeeder grid.")


# -----------------------------
# Autosave and recovery
# -----------------------------

def auto_save(username: str, project_slug: str, state: Dict[str, Any]) -> None:
    """Write autosave file to history_database/autosave/ as JSON. Overwrites existing autosave for same user+project.
    Also updates session_state['last_autosave_time'].
    """
    path = _autosave_path(username, project_slug)
    payload = {
        'saved_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'state': state,
    }
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        st.session_state['last_autosave_time'] = payload['saved_at']
    except Exception as e:
        st.error(f"Auto-save failed: {e}")


def load_autosave(username: str, project_slug: str) -> Optional[Dict[str, Any]]:
    path = _autosave_path(username, project_slug)
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            payload = json.load(f)
        return payload
    except Exception:
        return None


def remove_autosave(username: str, project_slug: str) -> None:
    path = _autosave_path(username, project_slug)
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


# -----------------------------
# Helper to build combined fase2_measurements dict from session_state
# -----------------------------

def collect_fase2_from_state(port_count: int = 8) -> Dict[str, Any]:
    data = {}
    data['IOR_1310'] = st.session_state.get('IOR_1310', DEFAULT_IOR_1310)
    data['IOR_1550'] = st.session_state.get('IOR_1550', DEFAULT_IOR_1550)
    data['splitter_grid'] = st.session_state.get('df_splitter_fdt', [])
    data['opm_distribution_grid'] = st.session_state.get('df_opm_distribution', [])
    # OTDR data depends on mode; both keys supported
    data['otdr_distribution_grid'] = st.session_state.get('df_otdr_cluster', []) or st.session_state.get('df_otdr_subfeeder', [])
    return data


# Exports
__all__ = [
    'render_ior_inputs',
    'render_modul_2a_splitter',
    'render_modul_2b_opm_dist',
    'render_modul_2c_2d_otdr',
    'auto_save',
    'load_autosave',
    'remove_autosave',
    'collect_fase2_from_state',
]
