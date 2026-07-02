"""
app.py

Streamlit orchestrator for Universal ATP Generator.
Implements UI Tab navigation, login stub, admin/DC scaffolding, Fase1/Fase2/Arsip tabs and integrates other modules.

This file follows the MASTER PROMPT contract and connects previously committed modules.
"""

from __future__ import annotations
import os
import io
import json
import time
import random
from datetime import datetime
from typing import Dict, Any

import streamlit as st

from system_config import DEFAULT_IOR_1310, DEFAULT_IOR_1550, normalize_mode, json_safe_filename
import parser_engine
import ui_technical_forms
import excel_injector_phase1
import excel_injector_phase2

# Ensure directories
HISTORY_DIR = "history_database"
AUTOSAVE_DIR = os.path.join(HISTORY_DIR, "autosave")
os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(AUTOSAVE_DIR, exist_ok=True)

TEMPLATE_PATH = "Template.xlsx"
GREETINGS_FILE = "greetings.txt"

# -----------------------------
# Greeting engine
# -----------------------------

def _load_greetings(path: str) -> Dict[str, list]:
    sections = {"PAGI": [], "SIANG": [], "SORE": [], "MALAM": [], "MINGGU": []}
    if not os.path.exists(path):
        return sections
    cur = None
    with open(path, 'r', encoding='utf-8') as f:
        for raw in f:
            line = raw.strip()
            if not line:
                continue
            if line.startswith('[') and line.endswith(']'):
                k = line.strip('[]')
                cur = k
                continue
            if cur and cur in sections:
                sections[cur].append(line)
    return sections


def _choose_greeting(sections: Dict[str, list], name: str) -> str:
    now = datetime.now()
    # Determine slot
    weekday = now.weekday()  # Monday 0 ... Sunday 6
    hour = now.hour
    if weekday == 6:
        slot = 'MINGGU'
    elif 0 <= hour <= 10:
        slot = 'PAGI'
    elif 11 <= hour <= 14:
        slot = 'SIANG'
    elif 15 <= hour <= 18:
        slot = 'SORE'
    else:
        slot = 'MALAM'
    choices = sections.get(slot) or sections.get('PAGI') or [f"Halo, {name}. Kerja dulu, istirahat belakangan."]
    text = random.choice(choices)
    return text.format(nama=name)

# load greetings once
GREETINGS = _load_greetings(GREETINGS_FILE)

# -----------------------------
# Helpers: JSON storage
# -----------------------------

def _project_filename(nama_lokasi: str) -> str:
    base = parser_engine.normalize_json_filename(nama_lokasi)
    return base


def save_project_json(metadata: Dict[str, Any], fase1_structures: Dict[str, Any], fase2_measurements: Dict[str, Any], user: str) -> str:
    nama = metadata.get('NAMA_LOKASI', 'untitled')
    fname = _project_filename(nama)
    path = os.path.join(HISTORY_DIR, fname)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    doc = {
        'saved_at': now,
        'project_mode': st.session_state.get('project_mode', 'cluster'),
        'created_by': st.session_state.get('created_by', user),
        'fase1_completed_at': st.session_state.get('fase1_completed_at'),
        'modified_by': user,
        'fase2_completed_at': st.session_state.get('fase2_completed_at'),
        'shared_with': st.session_state.get('shared_with', []),
        'metadata': metadata,
        'fase1_structures': fase1_structures,
        'fase2_measurements': fase2_measurements,
    }
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return path


def load_project_json(path: str) -> Dict[str, Any]:
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def list_projects() -> list:
    items = []
    for fn in os.listdir(HISTORY_DIR):
        if fn.endswith('.json'):
            items.append(fn)
    items.sort()
    return items

# -----------------------------
# Session init defaults
# -----------------------------

if 'is_logged_in' not in st.session_state:
    st.session_state['is_logged_in'] = False
    st.session_state['user_role'] = None
    st.session_state['user_name'] = None

# Metadata defaults per spec §13
DEFAULT_METADATA = {
    'NAMA_PROYEK': 'EMR FTTH PROJECT',
    'REGION': 'JAWA TIMUR',
    'NAMA_LOKASI': 'DUSUN BOGO RW 08 FDT-2',
    'ID_LOKASI': 'NJK000095',
    'ALAMAT': 'Nglawak Kecamatan Kertosono',
    'NAMA_OLT': 'KERTOSONO',
    'ID_FDT_FROM': 'NJK.100.021.DSBG08-FDT2.019.110',
    'ID_FAT_TO': 'DSBG08FDT2.019',
    'NAMA_PT_VENDOR': 'PT Buana Menara Indonesia',
    'REP_VENDOR': 'ERFIN FIRMANSYAH',
    'JABATAN_VENDOR': 'BMI FIELD SUPERVISOR',
    'NAMA_PT_CUSTOMER': 'PT Ekamas Mora Republik Tbk',
    'REP_CUSTOMER': 'M. NUGROHO',
    'JABATAN_CUSTOMER': 'EMR FIELD SUPERVISOR',
    'TANGGAL_TEST': '2026-06-27',
    'NO_PO': 'PO-EMR-2026-001',
}

if 'metadata' not in st.session_state:
    st.session_state['metadata'] = DEFAULT_METADATA.copy()

if 'fat_commands' not in st.session_state:
    st.session_state['fat_commands'] = []
if 'pole_commands' not in st.session_state:
    st.session_state['pole_commands'] = []
if 'parsed_fat' not in st.session_state:
    st.session_state['parsed_fat'] = []
if 'parsed_poles' not in st.session_state:
    st.session_state['parsed_poles'] = []
if 'fase1_extracted' not in st.session_state:
    st.session_state['fase1_extracted'] = False
if 'project_mode' not in st.session_state:
    st.session_state['project_mode'] = 'cluster'
if 'shared_with' not in st.session_state:
    st.session_state['shared_with'] = []
if 'last_autosave_time' not in st.session_state:
    st.session_state['last_autosave_time'] = None

# -----------------------------
# Layout & Login
# -----------------------------
st.set_page_config(page_title="Universal ATP Generator", layout='wide')
st.title("Universal ATP Generator")

# Check Template.xlsx
if not os.path.exists(TEMPLATE_PATH):
    st.error(f"Template.xlsx not found in repository root. Please place Template.xlsx in project root.")
    st.stop()

# Sidebar: Login or profile
with st.sidebar:
    st.markdown("### Profile")
    if not st.session_state['is_logged_in']:
        username = st.text_input('Username')
        password = st.text_input('Password', type='password')
        role = st.selectbox('Role', ['Document Control', 'Admin'])
        if st.button('Sign In'):
            # Simple stub: accept any username/password — store role
            st.session_state['is_logged_in'] = True
            st.session_state['user_name'] = username or 'dc_user'
            st.session_state['user_role'] = role
            st.session_state['created_by'] = st.session_state['user_name']
            st.experimental_rerun()
    else:
        name = st.session_state.get('user_name', 'dc')
        st.write(f"Signed in as **{name}**")
        # greeting
        greeting = _choose_greeting(GREETINGS, name)
        st.write(greeting)
        st.write('---')
        theme = st.radio('Theme', ['Light', 'Dark'], index=0, key='current_theme')
        st.write(f"Auto-save: {st.session_state.get('last_autosave_time')}")
        if st.button('Sign Out'):
            # clear session
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.experimental_rerun()

# If not logged in, show Login only
if not st.session_state['is_logged_in']:
    st.stop()

# Main content: tabs
tab1, tab2, tab3 = st.tabs(['FASE 1', 'FASE 2', 'ARSIP'])

# -----------------------------
# Tab 1: Fase 1
# -----------------------------
with tab1:
    st.header('FASE 1: Struktur Administrasi')
    col1, col2 = st.columns(2)
    with col1:
        # metadata form
        md = st.session_state['metadata']
        md['NAMA_PROYEK'] = st.text_input('NAMA_PROYEK', value=md.get('NAMA_PROYEK', ''), key='md_NAMA_PROYEK')
        md['REGION'] = st.text_input('REGION', value=md.get('REGION', ''), key='md_REGION')
        md['NAMA_LOKASI'] = st.text_input('NAMA_LOKASI', value=md.get('NAMA_LOKASI', ''), key='md_NAMA_LOKASI')
        md['ID_LOKASI'] = st.text_input('ID_LOKASI', value=md.get('ID_LOKASI', ''), key='md_ID_LOKASI')
        md['ALAMAT'] = st.text_input('ALAMAT', value=md.get('ALAMAT', ''), key='md_ALAMAT')
        md['NAMA_OLT'] = st.text_input('NAMA_OLT', value=md.get('NAMA_OLT', ''), key='md_NAMA_OLT')
        md['ID_FDT_FROM'] = st.text_input('ID_FDT_FROM', value=md.get('ID_FDT_FROM', ''), key='md_ID_FDT_FROM')
        md['ID_FAT_TO'] = st.text_input('ID_FAT_TO', value=md.get('ID_FAT_TO', ''), key='md_ID_FAT_TO')
    with col2:
        md['NAMA_PT_VENDOR'] = st.text_input('NAMA_PT_VENDOR', value=md.get('NAMA_PT_VENDOR', ''), key='md_NAMA_PT_VENDOR')
        md['REP_VENDOR'] = st.text_input('REP_VENDOR', value=md.get('REP_VENDOR', ''), key='md_REP_VENDOR')
        md['JABATAN_VENDOR'] = st.text_input('JABATAN_VENDOR', value=md.get('JABATAN_VENDOR', ''), key='md_JABATAN_VENDOR')
        md['NAMA_PT_CUSTOMER'] = st.text_input('NAMA_PT_CUSTOMER', value=md.get('NAMA_PT_CUSTOMER', ''), key='md_NAMA_PT_CUSTOMER')
        md['REP_CUSTOMER'] = st.text_input('REP_CUSTOMER', value=md.get('REP_CUSTOMER', ''), key='md_REP_CUSTOMER')
        md['JABATAN_CUSTOMER'] = st.text_input('JABATAN_CUSTOMER', value=md.get('JABATAN_CUSTOMER', ''), key='md_JABATAN_CUSTOMER')
        md['TANGGAL_TEST'] = st.date_input('TANGGAL_TEST', value=st.session_state['metadata'].get('TANGGAL_TEST'))
        md['NO_PO'] = st.text_input('NO_PO', value=md.get('NO_PO', ''), key='md_NO_PO')
    st.markdown('---')
    st.subheader('Komando FAT (contoh: A12 untuk FAT A01..A12)')
    fat_raw = st.text_area('Komando FAT', value='\n'.join(st.session_state.get('fat_commands', [])), height=80)
    st.subheader('Komando Tiang (contoh: pole 73=3, ext 74=2)')
    pole_raw = st.text_area('Komando Tiang', value='\n'.join(st.session_state.get('pole_commands', [])), height=80)

    if st.button('Ekstrak & Validasi Struktur'):
        # parse
        fat_lines = [l.strip() for l in fat_raw.splitlines() if l.strip()]
        pole_lines = [l.strip() for l in pole_raw.splitlines() if l.strip()]
        parsed_fat = parser_engine.parse_fat_inline(fat_lines)
        parsed_poles = parser_engine.parse_pole_inline(pole_lines)
        st.session_state['fat_commands'] = fat_lines
        st.session_state['pole_commands'] = pole_lines
        st.session_state['parsed_fat'] = parsed_fat
        st.session_state['parsed_poles'] = parsed_poles
        st.session_state['fase1_extracted'] = True
        st.session_state['fase1_completed_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        st.success(f"Ekstraksi selesai. {len(parsed_fat)} FAT dan {len(parsed_poles)} group tiang diekstrak.")

    st.markdown('---')
    cold1, cold2 = st.columns(2)
    with cold1:
        if st.button('Download Draf Fase 1'):
            # create draft bytes
            metadata = {k: v for k, v in st.session_state['metadata'].items()}
            mode = st.session_state.get('project_mode', 'cluster')
            bio = excel_injector_phase1.inject_excel_fase1_draft(TEMPLATE_PATH, metadata, st.session_state.get('parsed_fat', []), st.session_state.get('parsed_poles', []), mode)
            # filename: DRAFT_{MODE} {NAMA_LOKASI}.xlsx
            mode_code = 'CL' if mode == 'cluster' else 'SF'
            nama = metadata.get('NAMA_LOKASI', 'untitled')
            filename = f"DRAFT_{mode_code} {nama}.xlsx"
            st.download_button('Klik untuk unduh Draf Fase 1', data=bio, file_name=filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with cold2:
        if st.button('Amankan Project ke JSON'):
            if not st.session_state['fase1_extracted']:
                st.error('Ekstrak struktur Fase 1 terlebih dahulu.')
            else:
                metadata = {k: v for k, v in st.session_state['metadata'].items()}
                fase1_structures = {
                    'raw_fat_commands': st.session_state.get('fat_commands', []),
                    'raw_pole_commands': st.session_state.get('pole_commands', []),
                    'extracted_fat_list': st.session_state.get('parsed_fat', []),
                    'extracted_pole_list': st.session_state.get('parsed_poles', []),
                }
                fase2_measurements = {'IOR_1310': DEFAULT_IOR_1310, 'IOR_1550': DEFAULT_IOR_1550}
                save_path = save_project_json(metadata, fase1_structures, fase2_measurements, st.session_state['user_name'])
                st.success(f'Project diamankan ke {save_path}')

# -----------------------------
# Tab 2: Fase 2
# -----------------------------
with tab2:
    st.header('FASE 2: Data Teknis Lapangan')
    if not st.session_state['fase1_extracted']:
        st.info('Fase 1 belum diekstrak. Silakan kembali ke Tab FASE 1 dan tekan "Ekstrak & Validasi Struktur".')
    else:
        # Mode selection
        mode = st.radio('Pilih Mode Dokumen', ['Cluster Jaringan (Distribusi Hilir)', 'Subfeeder Jaringan (Backbone Hulu)'], index=0)
        mode_slug = 'cluster' if 'Cluster' in mode else 'subfeeder'
        st.session_state['project_mode'] = mode_slug

        # IOR inputs
        ui_technical_forms.render_ior_inputs()

        # Splitter
        ui_technical_forms.render_modul_2a_splitter(port_count=8)

        # OPM Distribution
        ui_technical_forms.render_modul_2b_opm_dist(st.session_state.get('parsed_fat', []), port_count=8)

        # OTDR
        ui_technical_forms.render_modul_2c_2d_otdr(st.session_state.get('parsed_fat', []), mode=mode_slug)

        st.markdown('---')
        # Auto-save controls
        colA, colB = st.columns([1,3])
        with colA:
            autosave_enable = st.checkbox('Aktifkan Auto-Save (manual trigger tiap perubahan)', value=False, key='autosave_enable')
            if autosave_enable:
                # perform immediate autosave
                payload_state = {
                    'metadata': st.session_state.get('metadata', {}),
                    'parsed_fat': st.session_state.get('parsed_fat', []),
                    'parsed_poles': st.session_state.get('parsed_poles', []),
                    'fase2_grids': ui_technical_forms.collect_fase2_from_state(),
                }
                ui_technical_forms.auto_save(st.session_state.get('user_name', 'anonymous'), st.session_state['metadata'].get('NAMA_LOKASI', 'untitled'), payload_state)
                st.success('Auto-save triggered')
        with colB:
            last = st.session_state.get('last_autosave_time')
            st.write(f"💾 Tersimpan otomatis: {last}")

        if st.button('Generate Dokumen Final (F1+F2)'):
            # assemble metadata + fase2 and call injector final
            metadata = st.session_state.get('metadata', {})
            parsed_fat = st.session_state.get('parsed_fat', [])
            parsed_poles = st.session_state.get('parsed_poles', [])
            phase2 = ui_technical_forms.collect_fase2_from_state()
            # store modified_by & timestamp in JSON later
            bio = excel_injector_phase2.inject_excel_final(TEMPLATE_PATH, metadata, parsed_fat, parsed_poles, phase2, mode_slug, port_count=8)
            mode_code = 'CL' if mode_slug == 'cluster' else 'SF'
            nama = metadata.get('NAMA_LOKASI', 'untitled')
            filename = f"DOC_{mode_code} {nama}.xlsx"
            st.download_button('Klik untuk unduh Dokumen Final', data=bio, file_name=filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            # update project JSON if exists and remove autosave
            # Save updated JSON
            fase1_structures = {
                'raw_fat_commands': st.session_state.get('fat_commands', []),
                'raw_pole_commands': st.session_state.get('pole_commands', []),
                'extracted_fat_list': st.session_state.get('parsed_fat', []),
                'extracted_pole_list': st.session_state.get('parsed_poles', []),
            }
            save_path = save_project_json(metadata, fase1_structures, phase2, st.session_state.get('user_name', 'anonymous'))
            st.session_state['last_autosave_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ui_technical_forms.remove_autosave(st.session_state.get('user_name', 'anonymous'), metadata.get('NAMA_LOKASI', 'untitled'))
            st.success(f'Dokumen final dibuat dan project JSON diperbarui: {save_path}')

# -----------------------------
# Tab 3: Arsip
# -----------------------------
with tab3:
    st.header('PUSAT ARSIP DIGITAL')
    projects = list_projects()
    q = st.text_input('Cari arsip (nama file atau lokasi)', '')
    filtered = [p for p in projects if q.lower() in p.lower()]
    st.write(f"Ditemukan {len(filtered)} arsip")
    for p in filtered:
        col1, col2, col3 = st.columns([3,1,1])
        col1.write(p)
        if col2.button('Muat Project', key=f'load_{p}'):
            path = os.path.join(HISTORY_DIR, p)
            doc = load_project_json(path)
            # load into session_state
            st.session_state['metadata'] = doc.get('metadata', DEFAULT_METADATA.copy())
            st.session_state['parsed_fat'] = doc.get('fase1_structures', {}).get('extracted_fat_list', [])
            st.session_state['parsed_poles'] = doc.get('fase1_structures', {}).get('extracted_pole_list', [])
            st.session_state['fase1_extracted'] = True
            st.success(f'Project {p} dimuat ke session.')
        if col3.button('Kelola Izin', key=f'share_{p}'):
            path = os.path.join(HISTORY_DIR, p)
            doc = load_project_json(path)
            owner = doc.get('created_by')
            if st.session_state.get('user_name') != owner and st.session_state.get('user_role') != 'Admin':
                st.error('Hanya pemilik proyek atau Admin yang dapat mengelola izin.')
            else:
                shared = doc.get('shared_with', [])
                new_shared = st.text_input('shared_with (koma pisah)', ','.join(shared))
                if st.button('Simpan Izin'):
                    arr = [s.strip() for s in new_shared.split(',') if s.strip()]
                    doc['shared_with'] = arr
                    with open(path, 'w', encoding='utf-8') as f:
                        json.dump(doc, f, ensure_ascii=False, indent=2)
                    st.success('Izin proyek diperbarui.')

# Footer
st.markdown('---')
st.markdown('<div style="text-align:center;color:gray">Developed by Senior Python/Streamlit Developer</div>', unsafe_allow_html=True)
