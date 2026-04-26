# -*- coding: utf-8 -*-
"""
Wochenplan Scheduler — Streamlit UI
Run: streamlit run streamlit_app.py
"""

import json
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

import wochenplan_scheduler as sched

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

STAFF_JSON  = Path(__file__).parent / "staff.json"
LAYOUT_JSON = Path(__file__).parent / "layout.json"
MEETING_POOLS_JSON = Path(__file__).parent / "meeting_pools.json"
TEMPLATE_XLSM = Path(__file__).parent / "KW_xx_TEMPLATE.xlsm"
OG_LIST_NO_LAUFEN = [og for og in sched.OG_LIST if og != "Laufen"]
ALL_WEEKDAYS = sched.WEEKDAYS  # ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]

# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------

def load_layout() -> dict:
    with open(LAYOUT_JSON, encoding="utf-8") as f:
        return json.load(f)


def save_layout(data: dict) -> None:
    with open(LAYOUT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    sched.load_layout_from_json(str(LAYOUT_JSON))


def _cells_to_str(cells) -> str:
    if isinstance(cells, (list, tuple)):
        return ", ".join(cells)
    return str(cells) if cells else ""


def _str_to_cells(s) -> list:
    if not isinstance(s, str) or not s.strip():
        return []
    return [c.strip() for c in s.replace(";", ",").split(",") if c.strip()]


# ---------------------------------------------------------------------------
# Staff persistence helpers
# ---------------------------------------------------------------------------

def save_staff_to_json() -> None:
    """Serialise current sched.staff_by_name to staff.json."""
    records = [
        {
            "name": s.name,
            "role": s.role,
            "site": s.site,
            "leads_ogs": sorted(s.leads_ogs),
            "rotations": sorted(s.rotations),
            "fr_excluded": s.fr_excluded,
            "fr_excluded_days": sorted(s.fr_excluded_days),
            "absent_by_default": s.absent_by_default,
            "covers_for": s.covers_for,
        }
        for s in sched.staff_by_name.values()
    ]
    with open(STAFF_JSON, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def staff_to_display_dataframe() -> pd.DataFrame:
    """Read-only overview table shown above the edit form."""
    rows = []
    # Sort by surname (last part of name after last space)
    def get_surname(s):
        return s.name.split()[-1]
    
    for s in sorted(sched.staff_by_name.values(), key=get_surname):
        fr_info = "Immer" if s.fr_excluded else (
            ", ".join(d[:2] for d in sched.WEEKDAYS if d in s.fr_excluded_days) or "—"
        )
        rows.append({
            "Name": s.name,
            "Rolle": s.role,
            "Standort": s.site,
            "Organgruppenleitung": ", ".join(sorted(s.leads_ogs)) or "—",
            "Rotationen": ", ".join(sorted(s.rotations)) or "—",
            "Kein Frontarzt": fr_info,
            "Stellvertretung": f"Für {s.covers_for}" if s.absent_by_default and s.covers_for else ("Standardmäßig absent" if s.absent_by_default else "—"),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Session-state bootstrap (runs once per browser session)
# ---------------------------------------------------------------------------

def _init_session_state() -> None:
    if "staff_loaded" not in st.session_state:
        # sched already loaded staff.json at import time if it existed;
        # nothing more to do — just mark the session as initialised.
        st.session_state["staff_loaded"] = True
    if "result_bytes" not in st.session_state:
        st.session_state["result_bytes"] = None
    if "result_filename" not in st.session_state:
        st.session_state["result_filename"] = "Wochenplan_FINAL.xlsm"


_init_session_state()

st.set_page_config(
    page_title="Wochenplan Scheduler",
    page_icon="📋",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Password gate
# ---------------------------------------------------------------------------

def _check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    st.markdown(
        """
        <style>
        .login-box {
            max-width: 340px;
            margin: 8rem auto 0 auto;
            padding: 2rem 2rem 1.5rem 2rem;
            border-radius: 8px;
            background: #1a1a1a;
            box-shadow: 0 4px 24px rgba(0,0,0,0.5);
            text-align: center;
        }
        .login-title {
            font-size: 1.2rem;
            font-weight: 600;
            color: #e0e0e0;
            margin-bottom: 1.5rem;
            letter-spacing: 0.05em;
        }
        </style>
        <div class="login-box">
            <div class="login-title">🔒 Wochenplan Scheduler</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            pw = st.text_input("Passwort", type="password", label_visibility="collapsed",
                               placeholder="Passwort eingeben")
            submitted = st.form_submit_button("Anmelden", use_container_width=True, type="primary")
            
        if submitted:
            if pw == st.secrets.get("password", ""):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Falsches Passwort.")
    return False

if not _check_password():
    st.stop()

st.title("Wochenplan Scheduler — KSBL Radiologie")
st.caption("by S. Vitéz · Powered by Anthropic")

# Helper functions for meeting pools (used in multiple tabs)
def load_meeting_pools() -> dict:
    with open(MEETING_POOLS_JSON, encoding="utf-8") as f:
        return json.load(f)

def save_meeting_pools(data: dict) -> None:
    with open(MEETING_POOLS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    sched.load_meeting_pools_from_json(str(MEETING_POOLS_JSON))


# ===========================================================================
# Helper Functions
# ===========================================================================

def _staff_form(form_key: str, defaults: dict | None = None) -> dict | None:
    """Render a staff edit/add form and return the submitted values as a dict,
    or None if the form was not submitted."""
    d = defaults or {}
    with st.form(form_key, clear_on_submit=(defaults is None)):
        c1, c2, c3 = st.columns(3)
        name     = c1.text_input("Name", value=d.get("name", ""), placeholder="J. Beispiel",
                                 disabled=(defaults is not None))
        role     = c2.selectbox("Rolle", ["AA", "OA", "LA"],
                                index=["AA", "OA", "LA"].index(d.get("role", "AA")))
        site     = c3.selectbox("Standort", ["BH", "LI"],
                                index=["BH", "LI"].index(d.get("site", "BH")))

        c4, c5 = st.columns(2)
        leads = c4.multiselect(
            "Organgruppenleitung",
            options=sched.OG_LIST,
            default=sorted(d.get("leads_ogs", [])),
            help="Nur Auswählen falls zutreffend. Laufen wird wie eine Organgruppenleitung gehandhabt",
        )
        rots = c5.multiselect(
            "Rotationen",
            options=OG_LIST_NO_LAUFEN,
            default=sorted(d.get("rotations", [])),
        )

        st.markdown("**Kein Frontarzt**")
        fr_col1, fr_col2 = st.columns([1, 2])
        # Only disable for AA when editing (defaults is not None)
        disable_for_aa = (role == "AA" and defaults is not None)
        fr_always = fr_col1.checkbox(
            "Nie Frontarzt",
            value=d.get("fr_excluded", False),
            disabled=disable_for_aa,
            help="Nur relevant für Fachärzte (OA/LA)." if disable_for_aa else None,
        )
        fr_days = fr_col2.multiselect(
            "Nur an diesen Tagen kein Frontarzt",
            options=sched.WEEKDAYS,
            default=sorted(d.get("fr_excluded_days", [])),
            disabled=(fr_always or disable_for_aa),
            help="Wird ignoriert wenn 'Nie Frontarzt' aktiviert ist oder Person kein Facharzt ist.",
        )

        st.markdown("**Stellvertretungsregel**")
        sub_col1, sub_col2 = st.columns([1, 2])
        absent_by_default = sub_col1.checkbox(
            "Standardmäßig absent",
            value=d.get("absent_by_default", False),
            help="Person ist grundsätzlich absent und erscheint nur wenn die vertretene Person fehlt.",
        )
        covers_for_options = [""] + sorted(sched.staff_by_name.keys())
        current_covers_for = d.get("covers_for") or ""
        covers_for_idx = covers_for_options.index(current_covers_for) if current_covers_for in covers_for_options else 0
        covers_for = sub_col2.selectbox(
            "Vertritt",
            options=covers_for_options,
            index=covers_for_idx,
            disabled=not absent_by_default,
            help="Person wird verfügbar wenn die ausgewählte Person absent ist.",
        )

        label = "Änderungen speichern" if defaults else "Hinzufügen"
        submitted = st.form_submit_button(label, type="primary")
        if submitted:
            return {
                "name": name.strip(),
                "role": role,
                "site": site,
                "leads_ogs": leads if role == "LA" else [],
                "rotations": rots,
                "fr_excluded": fr_always,
                "fr_excluded_days": [] if fr_always else fr_days,
                "absent_by_default": absent_by_default,
                "covers_for": covers_for if covers_for else None,
            }
    return None


tab_template, tab_eigene, tab_Feiertage, tab_personal, tab_pools, tab_laufen, tab_rapporte, tab_layout = st.tabs(["Wochenplan mit Standard-Vorlage", "Wochenplan mit eigener Vorlage", "Feiertage", "Personalverwaltung", "Rapporte-Pools", "Radiologe in Laufen", "Rapporte verwalten", "Layout-Editor"])

# ===========================================================================
# TAB 1 — Wochenplan mit Standard-Vorlage
# ===========================================================================

with tab_template:
    # Check if template exists
    template_exists = TEMPLATE_XLSM.exists()
    if not template_exists:
        st.warning("⚠️ Keine 'KW_xx_TEMPLATE.xlsm' gefunden. Bitte lade eine Vorlage hoch (unten).")
    
    st.markdown("### 📤 CSV + Standard-Vorlage 🪄→ fertiger Wochenplan")
    
    csv_file_opt1 = st.file_uploader(
        "CSV-Datei hochladen",
        type=["csv"],
        help="Lade nur eine CSV-Datei hoch. Die Standard-Vorlage wird automatisch geladen.",
        key="csv_opt1",
        disabled=not template_exists
    )
    
    col1, col2 = st.columns([1, 2])
    with col1:
        seed_opt1 = st.number_input(
            "Seed", value=1234, step=1, format="%d",
            help="Zufalls-Seed für reproduzierbare Ergebnisse.",
            key="seed_opt1"
        )
    
    run_opt1 = st.button(
        "Wochenplan erstellen (mit Standard-Vorlage)",
        disabled=(csv_file_opt1 is None or not template_exists),
        type="primary",
        key="run_opt1"
    )
    
    if run_opt1 and csv_file_opt1 and template_exists:
        csv_tmp_path = None
        output_tmp_path = None
        try:
            # Save CSV to temp
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as csv_tmp:
                csv_tmp.write(csv_file_opt1.getbuffer())
                csv_tmp_path = csv_tmp.name
            
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f_out:
                output_tmp_path = f_out.name
            
            with st.spinner("Pipeline läuft…"):
                # Configure Laufen days
                laufen_days = st.session_state.get("laufen_days", ["Dienstag"])
                sched.LAUFEN_DAYS.clear()
                sched.LAUFEN_DAYS.update(laufen_days)
                
                # Reset counters
                sched.reset_all_counters()
                
                # Load template from disk
                wb = load_workbook(str(TEMPLATE_XLSM), data_only=False, keep_vba=True)
                ws = wb["Wochenplan"]
                
                # Stage 0: Cleanup
                sched.cleanup_blocks(ws, clear_fr=True, clear_og=True, clear_meetings=True)
                
                # Stage 0.5: CSV import
                sched.fill_dienste_from_csv(ws, csv_tmp_path)
                
                # Read absences
                absences = sched.read_absences_by_day(ws)
                
                # Stage 1: OG
                sched.assign_la_to_ogs(ws, absences)
                sched.assign_nonleaders_to_ogs(ws, absences, seed=seed_opt1)
                
                # Stage 2: FR
                sched.assign_fr_shifts_to_cells(ws, absences, seed=seed_opt1)
                
                # Stage 3: Meetings
                sched.assign_meetings(ws, absences, seed=seed_opt1)
                
                # Save
                wb.save(output_tmp_path)
                wb.close()
            
            # Offer download
            with open(output_tmp_path, "rb") as f:
                st.download_button(
                    "⬇️ Finaler Wochenplan herunterladen",
                    data=f.read(),
                    file_name="Wochenplan_FINAL.xlsm",
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                )
            
            st.success("✔ Pipeline erfolgreich abgeschlossen!")
            sched.print_weekly_stats()
            
        except Exception as e:
            st.error(f"Fehler: {e}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            if csv_tmp_path and os.path.exists(csv_tmp_path):
                os.unlink(csv_tmp_path)
            if output_tmp_path and os.path.exists(output_tmp_path):
                os.unlink(output_tmp_path)
    

    st.divider()
    
    st.markdown("### 🗂️ Vorlage verwalten")
    st.caption("Lade eine neue leere Vorlage hoch, um die Standard-Vorlage zu ersetzen.")
    
    with st.expander("Neue Vorlage hochladen"):
        
        new_template = st.file_uploader(
            "Neue Vorlage hochladen",
            type=["xlsm"],
            help="Die neue leere Wochenplan-Vorlage. Wird automatisch als 'KW_xx_TEMPLATE.xlsm' gespeichert.",
            key="new_template_uploader"
        )
        
        if new_template:
            if st.button("Vorlage ersetzen", type="primary", key="replace_template_btn"):
                try:
                    with open(TEMPLATE_XLSM, "wb") as f:
                        f.write(new_template.getbuffer())
                    st.success(f"✓ '{new_template.name}' wurde als 'KW_xx_TEMPLATE.xlsm' gespeichert!")
                except Exception as e:
                    st.error(f"Fehler beim Ersetzen: {e}")

    
    
# ===========================================================================
# TAB 2 — Wochenplan mit eigener Vorlage
# ===========================================================================
 
with tab_eigene:

    st.markdown("### 📤 CSV + 📤 Eigene Vorlage 🪄→ fertiger Wochenplan")
    
    csv_file_opt2 = st.file_uploader(
        "CSV-Datei hochladen",
        type=["csv"],
        help="CSV-Datei mit Absenzen und Diensten.",
        key="csv_opt2"
    )
    
    template_file_opt2 = st.file_uploader(
        "Eigene Wochenplan-Vorlage (.xlsm) hochladen",
        type=["xlsm"],
        help="Lade eine eigene .xlsm-Vorlage hoch.",
        key="template_opt2"
    )
    
    col1, col2 = st.columns([1, 2])
    with col1:
        seed_opt2 = st.number_input(
            "Seed", value=1234, step=1, format="%d",
            help="Zufalls-Seed für reproduzierbare Ergebnisse.",
            key="seed_opt2"
        )
    
    run_opt2 = st.button(
        "Wochenplan erstellen (mit eigener Vorlage)",
        disabled=(csv_file_opt2 is None or template_file_opt2 is None),
        type="primary",
        key="run_opt2"
    )
    
    if run_opt2 and csv_file_opt2 and template_file_opt2:
        csv_tmp_path = None
        template_tmp_path = None
        output_tmp_path = None
        try:
            # Save files to temp
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as csv_tmp:
                csv_tmp.write(csv_file_opt2.getbuffer())
                csv_tmp_path = csv_tmp.name
            
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmpl_tmp:
                tmpl_tmp.write(template_file_opt2.getbuffer())
                template_tmp_path = tmpl_tmp.name
            
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f_out:
                output_tmp_path = f_out.name
            
            with st.spinner("Pipeline läuft…"):
                # Configure Laufen days
                laufen_days = st.session_state.get("laufen_days", ["Dienstag"])
                sched.LAUFEN_DAYS.clear()
                sched.LAUFEN_DAYS.update(laufen_days)
                
                # Reset counters
                sched.reset_all_counters()
                
                # Load uploaded template
                wb = load_workbook(template_tmp_path, data_only=False, keep_vba=True)
                ws = wb["Wochenplan"]
                
                # Stage 0: Cleanup
                sched.cleanup_blocks(ws, clear_fr=True, clear_og=True, clear_meetings=True)
                
                # Stage 0.5: CSV import
                sched.fill_dienste_from_csv(ws, csv_tmp_path)
                
                # Read absences
                absences = sched.read_absences_by_day(ws)
                
                # Stage 1: OG
                sched.assign_la_to_ogs(ws, absences)
                sched.assign_nonleaders_to_ogs(ws, absences, seed=seed_opt2)
                
                # Stage 2: FR
                sched.assign_fr_shifts_to_cells(ws, absences, seed=seed_opt2)
                
                # Stage 3: Meetings
                sched.assign_meetings(ws, absences, seed=seed_opt2)
                
                # Save
                wb.save(output_tmp_path)
                wb.close()
            
            # Offer download
            with open(output_tmp_path, "rb") as f:
                st.download_button(
                    "⬇️ Finaler Wochenplan herunterladen",
                    data=f.read(),
                    file_name="Wochenplan_FINAL.xlsm",
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                )
            
            st.success("✔ Pipeline erfolgreich abgeschlossen!")
            sched.print_weekly_stats()
            
        except Exception as e:
            st.error(f"Fehler: {e}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            if csv_tmp_path and os.path.exists(csv_tmp_path):
                os.unlink(csv_tmp_path)
            if template_tmp_path and os.path.exists(template_tmp_path):
                os.unlink(template_tmp_path)
            if output_tmp_path and os.path.exists(output_tmp_path):
                os.unlink(output_tmp_path)

# ===========================================================================
# TAB 3 — Feiertage
# ===========================================================================
with tab_Feiertage:

    st.markdown("### 🎉 Feiertage")
    st.caption("Wähle Wochentage, die als Feiertage behandelt werden sollen (keine Absenzen, OG, FR, Rapporte).")
    
    # Load current feiertage from layout
    layout = load_layout()
    current_feiertage = layout.get("feiertage", [])
    
    feiertage_selected = st.multiselect(
        "Feiertage",
        options=["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"],
        default=current_feiertage,
        help="An diesen Tagen werden nur Nacht-, Hintergrund- und Vordergrunddienste eingeplant.",
        key="feiertage_select"
    )
    
    if st.button("Feiertage speichern", type="primary", key="save_feiertage_btn"):
        layout["feiertage"] = feiertage_selected
        save_layout(layout)
        st.success(f"✓ Feiertage gespeichert: {', '.join(feiertage_selected) if feiertage_selected else 'Keine'}")
        st.rerun()

# ===========================================================================
# TAB 4 — Personalverwaltung
# ===========================================================================
with tab_personal:
    st.subheader("Personalbestand")

    # Read-only overview table with color-coded roles
    df = staff_to_display_dataframe()
    
    # Define role colors (softer, 30% opacity)
    def color_row(row):
        colors = {
            "AA": "background-color: rgba(144, 238, 144, 0.3)",  # light green
            "OA": "background-color: rgba(135, 206, 235, 0.3)",  # sky blue
            "LA": "background-color: rgba(255, 182, 198, 0.3)",  # light pink
        }
        color = colors.get(row["Rolle"], "")
        return [color] * len(row)
    
    # Apply styling to entire rows
    styled_df = df.style.apply(color_row, axis=1)
    
    st.dataframe(
        styled_df,
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    # ---- Edit existing staff member ----
    st.subheader("Mitarbeiter bearbeiten")
    all_names = sorted(sched.staff_by_name.keys())
    edit_name = st.selectbox("Person auswählen", all_names, key="edit_select")

    if edit_name:
        s = sched.staff_by_name[edit_name]
        edit_defaults = {
            "name": s.name,
            "role": s.role,
            "site": s.site,
            "leads_ogs": sorted(s.leads_ogs),
            "rotations": sorted(s.rotations),
            "fr_excluded": s.fr_excluded,
            "fr_excluded_days": sorted(s.fr_excluded_days),
            "absent_by_default": s.absent_by_default,
            "covers_for": s.covers_for,
        }
        result = _staff_form("edit_staff_form", defaults=edit_defaults)
        if result:
            sched.add_staff(
                name=edit_name,
                role=result["role"],
                site=result["site"],
                leads_for=result["leads_ogs"],
                rotation=result["rotations"],
                fr_excluded=result["fr_excluded"],
                fr_excluded_days=result["fr_excluded_days"],
                absent_by_default=result["absent_by_default"],
                covers_for=result["covers_for"],
            )
            sched.rebuild_quick_views()
            save_staff_to_json()
            st.success(f"'{edit_name}' aktualisiert und gespeichert.")
            st.rerun()

        if st.button("Mitarbeiter löschen", key="delete_btn"):
            deleted_name = edit_name
            del sched.staff_by_name[deleted_name]
            sched.rebuild_quick_views()
            save_staff_to_json()
            
            # Clean up orphaned names from meeting_pools.json
            pools_data = load_meeting_pools()
            pools_modified = False
            all_current_names = set(sched.staff_by_name.keys())
            
            for meeting_key, cfg in pools_data.items():
                for pool in cfg.get("pools", []):
                    # Clean names list
                    if "names" in pool and pool["names"]:
                        original_names = pool["names"]
                        pool["names"] = [n for n in original_names if n in all_current_names]
                        if len(pool["names"]) != len(original_names):
                            pools_modified = True
                    
                    # Clean exclude_names
                    if "exclude_names" in pool and pool["exclude_names"]:
                        original_excluded = pool["exclude_names"]
                        pool["exclude_names"] = [n for n in original_excluded if n in all_current_names]
                        if len(pool["exclude_names"]) != len(original_excluded):
                            pools_modified = True
                        if not pool["exclude_names"]:
                            pool["exclude_names"] = None
                    
                    # Clean exclude_if_day
                    if "exclude_if_day" in pool and pool["exclude_if_day"]:
                        for day, names in list(pool["exclude_if_day"].items()):
                            cleaned = [n for n in names if n in all_current_names]
                            if cleaned:
                                pool["exclude_if_day"][day] = cleaned
                            else:
                                del pool["exclude_if_day"][day]
                            if cleaned != names:
                                pools_modified = True
                        if not pool["exclude_if_day"]:
                            pool["exclude_if_day"] = None
            
            if pools_modified:
                save_meeting_pools(pools_data)
            
            st.success(f"'{deleted_name}' wurde entfernt" + (" (und aus Rapporte-Pools entfernt)." if pools_modified else "."))
            st.rerun()

    st.divider()

    # ---- Add new staff member ----
    with st.expander("Neues Personalmitglied hinzufügen"):
        result = _staff_form("add_staff_form")
        if result:
            name_clean = result["name"]
            if not name_clean:
                st.warning("Name darf nicht leer sein.")
            elif name_clean in sched.staff_by_name:
                st.warning(f"'{name_clean}' ist bereits im Personalbestand.")
            else:
                sched.add_staff(
                    name=name_clean,
                    role=result["role"],
                    site=result["site"],
                    leads_for=result["leads_ogs"],
                    rotation=result["rotations"],
                    fr_excluded=result["fr_excluded"],
                    fr_excluded_days=result["fr_excluded_days"],
                    absent_by_default=result["absent_by_default"],
                    covers_for=result["covers_for"],
                )
                sched.rebuild_quick_views()
                save_staff_to_json()
                st.success(f"'{name_clean}' wurde hinzugefügt und gespeichert.")
                st.rerun()

# ===========================================================================
# TAB 5 — Layout-Editor
# ===========================================================================

_COL_CFG = {
    day: st.column_config.TextColumn(day[:2], help=day)
    for day in ALL_WEEKDAYS
}



# ===========================================================================
# TAB 7 — Rapporte verwalten
# ===========================================================================

def _rapport_overview_df() -> pd.DataFrame:
    """Overview table of all rapporte from meeting_pools.json."""
    pools_data = load_meeting_pools()
    rows = []
    for key in pools_data:
        parts = key.split("|", 1)
        site = parts[0] if len(parts) == 2 else "?"
        name = parts[1] if len(parts) == 2 else key
        rows.append({"Key": key, "Name": name, "Standort": site})
    return pd.DataFrame(rows)


def _rename_rapport(old_key: str, new_key: str) -> None:
    """Rename a rapport key atomically in both layout.json and meeting_pools.json."""
    # --- meeting_pools.json ---
    pools_data = load_meeting_pools()
    if old_key not in pools_data:
        raise KeyError(f"Rapport '{old_key}' nicht in meeting_pools.json gefunden.")
    cfg = pools_data.pop(old_key)
    pools_data[new_key] = cfg
    save_meeting_pools(pools_data)

    # --- layout.json ---
    layout = load_layout()
    old_parts = old_key.split("|", 1)
    new_parts = new_key.split("|", 1)
    old_site, old_name = old_parts[0], old_parts[1]
    new_site, new_name = new_parts[0], new_parts[1]

    # Move cell mappings: remove from old site/name, insert under new site/name
    old_cells = layout["meeting_cells"].get(old_site, {}).pop(old_name, {})
    layout["meeting_cells"].setdefault(new_site, {})[new_name] = old_cells
    save_layout(layout)


def _delete_rapport(key: str) -> None:
    """Delete a rapport from both layout.json and meeting_pools.json."""
    pools_data = load_meeting_pools()
    pools_data.pop(key, None)
    save_meeting_pools(pools_data)

    layout = load_layout()
    parts = key.split("|", 1)
    if len(parts) == 2:
        layout["meeting_cells"].get(parts[0], {}).pop(parts[1], None)
    save_layout(layout)


def _add_rapport(key: str, exclude_laufen: bool = False) -> None:
    """Add a new rapport to both layout.json and meeting_pools.json with empty defaults."""
    pools_data = load_meeting_pools()
    pools_data[key] = {
        "site": key.split("|", 1)[0],
        "pools": [{"type": "names", "names": [], "site": key.split("|", 1)[0], "exclude_laufen": exclude_laufen}],
        "fallback_text": "FÄLLT AUS",
        "roter_fallback_text": True,
    }
    save_meeting_pools(pools_data)

    layout = load_layout()
    parts = key.split("|", 1)
    site, name = parts[0], parts[1]
    layout["meeting_cells"].setdefault(site, {}).setdefault(name, {
        day: [] for day in ALL_WEEKDAYS
    })
    save_layout(layout)


with tab_rapporte:
    st.subheader("Rapporte verwalten")
    st.caption("Rapporte hinzufügen, umbenennen oder löschen. Zellen werden im Layout-Editor bearbeitet.")

    # Overview table
    rapport_df = _rapport_overview_df()
    st.dataframe(rapport_df[["Name", "Standort"]], use_container_width=True, hide_index=True)

    st.divider()

    # ---- Edit existing rapport ----
    st.subheader("Rapport bearbeiten")
    all_rapport_keys = rapport_df["Key"].tolist() if not rapport_df.empty else []
    edit_rapport_key = st.selectbox("Rapport auswählen", all_rapport_keys, key="edit_rapport_select")

    if edit_rapport_key:
        r_parts = edit_rapport_key.split("|", 1)
        r_site_current = r_parts[0]
        r_name_current = r_parts[1] if len(r_parts) == 2 else edit_rapport_key

        rc1, rc2 = st.columns(2)
        r_name_new = rc1.text_input("Name", value=r_name_current, key="edit_rapport_name")
        r_site_new = rc2.selectbox("Standort", ["BH", "LI"],
                                   index=["BH", "LI"].index(r_site_current) if r_site_current in ["BH", "LI"] else 0,
                                   key="edit_rapport_site")

        if st.button("Änderungen speichern", type="primary", key="save_rapport_btn"):
            new_key = f"{r_site_new}|{r_name_new.strip()}"
            if not r_name_new.strip():
                st.warning("Name darf nicht leer sein.")
            elif new_key != edit_rapport_key and new_key in all_rapport_keys:
                st.warning(f"'{new_key}' existiert bereits.")
            else:
                try:
                    _rename_rapport(edit_rapport_key, new_key)
                    st.success(f"Rapport umbenannt zu '{new_key}'.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Fehler: {e}")

        if st.button("Rapport löschen", key="delete_rapport_btn"):
            try:
                _delete_rapport(edit_rapport_key)
                st.success(f"'{edit_rapport_key}' wurde gelöscht.")
                st.rerun()
            except Exception as e:
                st.error(f"Fehler: {e}")

    st.divider()

    # ---- Add new rapport ----
    with st.expander("Neuen Rapport hinzufügen"):
        nc1, nc2 = st.columns(2)
        new_r_name = nc1.text_input("Name", placeholder="z.B. Medizin (07:45-08:00)", key="new_rapport_name")
        new_r_site = nc2.selectbox("Standort", ["BH", "LI"], key="new_rapport_site")

        if st.button("Rapport hinzufügen", type="primary", key="add_rapport_btn"):
            new_r_key = f"{new_r_site}|{new_r_name.strip()}"
            if not new_r_name.strip():
                st.warning("Name darf nicht leer sein.")
            elif new_r_key in all_rapport_keys:
                st.warning(f"'{new_r_key}' existiert bereits.")
            else:
                try:
                    _add_rapport(new_r_key, exclude_laufen=st.session_state.get("global_exclude_laufen", False))
                    st.success(f"'{new_r_key}' wurde hinzugefügt. Zellen im Layout-Editor eintragen.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Fehler: {e}")


with tab_layout:
    st.subheader("Layout-Editor")
    st.caption(
        "Excel-Zellreferenzen für alle Planabschnitte bearbeiten. "
        "Mehrere Zellen kommagetrennt eingeben, z.B. «T35, T36, T37»."
    )

    layout = load_layout()

    # ---- Abwesenheiten & Nachtdienst ----
    with st.expander("Abwesenheiten & Nachtdienst-Bereiche"):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Abwesenheiten** (Bereich je Tag)")
            abw_df = pd.DataFrame(
                [{"Tag": d, "Bereich": layout["abw_ranges"].get(d, "")} for d in ALL_WEEKDAYS]
            ).set_index("Tag")
            abw_edited = st.data_editor(abw_df, key="abw_ed", use_container_width=True)
        with c2:
            st.markdown("**Nachtdienst** (Bereich je Tag)")
            nacht_df = pd.DataFrame(
                [{"Tag": d, "Bereich": layout["nacht_ranges"].get(d, "")} for d in ALL_WEEKDAYS]
            ).set_index("Tag")
            nacht_edited = st.data_editor(nacht_df, key="nacht_ed", use_container_width=True)

    # ---- Spätdienst ----
    with st.expander("Spätdienst-Zellen"):
        spaet_rows = []
        for site in ["BH", "LI"]:
            row = {"Standort": site}
            for day in ALL_WEEKDAYS:
                row[day] = layout["spaetdienst_cells"].get(site, {}).get(day, "")
            spaet_rows.append(row)
        spaet_df = pd.DataFrame(spaet_rows).set_index("Standort")
        spaet_edited = st.data_editor(
            spaet_df, key="spaet_ed", use_container_width=True, column_config=_COL_CFG
        )

    # ---- Frontarzt ----
    with st.expander("Frontarzt-Zellen"):
        fr_rows = []
        for site in ["BH", "LI"]:
            row = {"Standort": site}
            for day in ALL_WEEKDAYS:
                row[day] = _cells_to_str(layout["fr_cells"].get(site, {}).get(day, []))
            fr_rows.append(row)
        fr_df = pd.DataFrame(fr_rows).set_index("Standort")
        fr_edited = st.data_editor(
            fr_df, key="fr_ed", use_container_width=True, column_config=_COL_CFG
        )

    # ---- Organgruppen ----
    with st.expander("Organgruppen-Zellen"):
        og_rows = []
        for og in sched.OG_LIST:
            row = {"OG": og}
            for day in ALL_WEEKDAYS:
                row[day] = _cells_to_str(layout["og_cells"].get(og, {}).get(day, []))
            og_rows.append(row)
        og_df = pd.DataFrame(og_rows).set_index("OG")
        og_edited = st.data_editor(
            og_df, key="og_ed", use_container_width=True, column_config=_COL_CFG
        )

    # ---- Rapporte BH ----
    with st.expander("Rapporte-Zellen — BH"):
        bh_rows = []
        for mtg, day_map in layout["meeting_cells"].get("BH", {}).items():
            row = {"Rapport": mtg}
            for day in ALL_WEEKDAYS:
                row[day] = _cells_to_str(day_map.get(day, []))
            bh_rows.append(row)
        bh_df = pd.DataFrame(bh_rows).set_index("Rapport")
        bh_edited = st.data_editor(
            bh_df, key="bh_mtg_ed", use_container_width=True, column_config=_COL_CFG
        )

    # ---- Rapporte LI ----
    with st.expander("Rapporte-Zellen — LI"):
        li_rows = []
        for mtg, day_map in layout["meeting_cells"].get("LI", {}).items():
            row = {"Rapport": mtg}
            for day in ALL_WEEKDAYS:
                row[day] = _cells_to_str(day_map.get(day, []))
            li_rows.append(row)
        li_df = pd.DataFrame(li_rows).set_index("Rapport")
        li_edited = st.data_editor(
            li_df, key="li_mtg_ed", use_container_width=True, column_config=_COL_CFG
        )

    # ---- Medizin Montag placeholder ----
    with st.expander("Medizin Montag Platzhalter"):
        st.caption("Zelle für «BITTE EINTRAGEN» am Montag je Standort")
        _mc = layout.get("medizin_monday_cells", {})
        mc1, mc2 = st.columns(2)
        medizin_bh = mc1.text_input("BH", value=_mc.get("BH", ""), key="medizin_bh")
        medizin_li = mc2.text_input("LI", value=_mc.get("LI", ""), key="medizin_li")

    # ---- Feiertage Merge Cells ----
    with st.expander("Zusammenführen von Zellen bei Feiertagen"):
        st.caption("Zellbereiche für Vordergrunddienst an Feiertagen (z.B. T23:AC24)")
        _fmc = layout.get("feiertage_merge_cells", {})
        
        fmc_rows = []
        for day in ALL_WEEKDAYS:
            fmc_rows.append({"Tag": day, "Bereich": _fmc.get(day, "")})
        
        fmc_df = pd.DataFrame(fmc_rows).set_index("Tag")
        feiertage_merge_edited = st.data_editor(
            fmc_df, 
            key="feiertage_merge_ed", 
            use_container_width=True,
            height=250
        )

    # ---- Save ----
    st.divider()
    if st.button("Alle Änderungen speichern", type="primary", key="save_layout_btn"):
        # Load current layout to preserve fields not in the editor
        current_layout = load_layout()
        
        new_layout = {
            "abw_ranges": {
                row: abw_edited.at[row, "Bereich"] for row in abw_edited.index
            },
            "nacht_ranges": {
                # Save edited weekdays (Mon-Fri)
                **{row: nacht_edited.at[row, "Bereich"] for row in nacht_edited.index},
                # Preserve weekend from current layout
                "Samstag": current_layout.get("nacht_ranges", {}).get("Samstag", ""),
                "Sonntag": current_layout.get("nacht_ranges", {}).get("Sonntag", ""),
            },
            "spaetdienst_cells": {
                site: {day: str(spaet_edited.at[site, day] or "") for day in ALL_WEEKDAYS}
                for site in ["BH", "LI"]
            },
            "fr_cells": {
                site: {day: _str_to_cells(fr_edited.at[site, day]) for day in ALL_WEEKDAYS}
                for site in ["BH", "LI"]
            },
            "og_cells": {
                og: {day: _str_to_cells(og_edited.at[og, day]) for day in ALL_WEEKDAYS}
                for og in sched.OG_LIST
            },
            "meeting_cells": {
                "BH": {
                    mtg: {day: _str_to_cells(bh_edited.at[mtg, day]) for day in ALL_WEEKDAYS}
                    for mtg in bh_edited.index
                },
                "LI": {
                    mtg: {day: _str_to_cells(li_edited.at[mtg, day]) for day in ALL_WEEKDAYS}
                    for mtg in li_edited.index
                },
            },
            "medizin_monday_cells": {
                "BH": medizin_bh.strip(),
                "LI": medizin_li.strip(),
            },
            "feiertage_merge_cells": {
                row: feiertage_merge_edited.at[row, "Bereich"] for row in feiertage_merge_edited.index
            },
            # Preserve new CSV import fields
            "vordergrunddienst_cells": current_layout.get("vordergrunddienst_cells", {}),
            "hintergrunddienst_cells": current_layout.get("hintergrunddienst_cells", {}),
            "date_cells": current_layout.get("date_cells", {}),
            "weekday_date_cells": current_layout.get("weekday_date_cells", {}),
            "feiertage": current_layout.get("feiertage", []),
        }
        save_layout(new_layout)
        st.success("Layout gespeichert und neu geladen.")
        st.rerun()

# ===========================================================================
# TAB 6 — Rapporte-Pools
# ===========================================================================

# Pool type options with display labels
_POOL_TYPES_MAP = {
    "names": "Person",
    "group": "Gruppe",
    "spaetdienst_aa": "Spätdienst_AA"
}
_POOL_TYPES = list(_POOL_TYPES_MAP.keys())
_POOL_TYPES_DISPLAY = list(_POOL_TYPES_MAP.values())

# Group options with display labels
_GROUP_MAP = {
    "AA": "AA",
    "OA": "OA",
    "LA": "LA",
    "FA_ALL": "alle Fachärzte"
}
_GROUP_OPTIONS = list(_GROUP_MAP.keys())
_GROUP_DISPLAY = list(_GROUP_MAP.values())

_SITE_OPTIONS = ["BH", "LI"]


def _exclude_if_day_to_str(eid: dict | None) -> str:
    """Convert {"Donnerstag": ["H.W. Ott"]} → 'Donnerstag: H.W. Ott'"""
    if not eid:
        return ""
    parts = []
    for day, names in eid.items():
        if isinstance(names, (list, tuple)):
            parts.append(f"{day}: {', '.join(names)}")
        else:
            parts.append(f"{day}: {names}")
    return "; ".join(parts)


def _str_to_exclude_if_day(s: str) -> dict | None:
    """Convert 'Donnerstag: H.W. Ott, X; Freitag: Y' → dict."""
    if not s or not s.strip():
        return None
    result = {}
    for part in s.split(";"):
        part = part.strip()
        if ":" not in part:
            continue
        day, names_str = part.split(":", 1)
        day = day.strip()
        names = [n.strip() for n in names_str.split(",") if n.strip()]
        if day and names:
            result[day] = names
    return result or None


def _list_to_str(lst: list | None) -> str:
    if not lst:
        return ""
    return ", ".join(str(x) for x in lst)


def _str_to_list(s: str) -> list:
    if not s or not s.strip():
        return []
    return [x.strip() for x in s.split(",") if x.strip()]


with tab_pools:
    st.subheader("Rapporte-Pools")
    st.caption(
        "Hier können die Prioritäts-Pools für jeden Rapport bearbeitet werden. "
        "Die Pools werden in der definierten Reihenfolge durchlaufen, bis eine "
        "verfügbare Person gefunden wird."
    )
    st.markdown("**Pools** (in Prioritätsreihenfolge)")

    pools_data = load_meeting_pools()

    for meeting_key, cfg in pools_data.items():
        with st.expander(meeting_key):
            prefix = meeting_key.replace("|", "_").replace(" ", "_").replace(":", "").replace("/", "_").replace("(", "").replace(")", "")

            pools = cfg.get("pools", [])

            # Get all staff names for dropdowns
            all_staff_names = sorted(sched.staff_by_name.keys())

            for i, pool in enumerate(pools):
                st.markdown(
                    f"<span style='font-size:1.05rem; font-weight:700; color:#4A90D9;'>Pool {i+1}</span>",
                    unsafe_allow_html=True,
                )
                st.markdown("---")
                pc1, pc2, pc3 = st.columns(3)

                # Typ dropdown with display labels
                current_type = pool.get("type", "names")
                type_display_idx = _POOL_TYPES.index(current_type) if current_type in _POOL_TYPES else 0
                pool_type_display = pc1.selectbox(
                    "Typ", 
                    options=_POOL_TYPES_DISPLAY,
                    index=type_display_idx,
                    key=f"{prefix}_p{i}_type",
                )
                pool_type = _POOL_TYPES[_POOL_TYPES_DISPLAY.index(pool_type_display)]
                pool["type"] = pool_type

                pool_site = pc2.selectbox(
                    "Standort", options=_SITE_OPTIONS,
                    index=_SITE_OPTIONS.index(pool.get("site", cfg.get("site", "BH"))),
                    key=f"{prefix}_p{i}_site",
                )
                pool["site"] = pool_site

                # Roter Text checkbox
                roter_text = pc3.checkbox(
                    "Roter Text",
                    value=(pool.get("style") == "red_bold"),
                    key=f"{prefix}_p{i}_rot",
                )
                pool["style"] = "red_bold" if roter_text else None

                # Type-specific fields
                if pool_type == "names":
                    current_names = pool.get("names") or []
                    selected_names = st.multiselect(
                        "Namen",
                        options=all_staff_names,
                        default=[n for n in current_names if n in all_staff_names],
                        key=f"{prefix}_p{i}_names",
                        help="Wählen Sie Personen für diesen Pool aus.",
                    )
                    pool["names"] = selected_names if selected_names else []

                if pool_type == "group":
                    current_group = pool.get("group", "AA")
                    group_display_idx = _GROUP_OPTIONS.index(current_group) if current_group in _GROUP_OPTIONS else 0
                    pool_group_display = st.selectbox(
                        "Gruppe", 
                        options=_GROUP_DISPLAY,
                        index=group_display_idx,
                        key=f"{prefix}_p{i}_group",
                    )
                    pool_group = _GROUP_OPTIONS[_GROUP_DISPLAY.index(pool_group_display)]
                    pool["group"] = pool_group

                # Spätdienst checkbox (single column now, Laufen moved to global)
                excl_spaet = st.checkbox(
                    "Spätdienst ausschließen",
                    value=bool(pool.get("exclude_spaetdienst")),
                    key=f"{prefix}_p{i}_excl_spaet",
                    help=f"Schließt Spätdienst-Personal von {pool_site} aus.",
                )
                pool["exclude_spaetdienst"] = pool_site if excl_spaet else None

                # Ausgeschlossene Personen
                current_excluded = pool.get("exclude_names") or []
                excluded_names = st.multiselect(
                    "Ausgeschlossene Personen",
                    options=all_staff_names,
                    default=[n for n in current_excluded if n in all_staff_names],
                    key=f"{prefix}_p{i}_excl_names",
                    help="Personen, die von diesem Pool ausgeschlossen sind.",
                )
                pool["exclude_names"] = excluded_names if excluded_names else None

                # Exclude if day
                eid_str = st.text_input(
                    "Ausschluss pro Tag",
                    value=_exclude_if_day_to_str(pool.get("exclude_if_day")),
                    key=f"{prefix}_p{i}_eid",
                    help="Format: 'Donnerstag: Name1, Name2; Freitag: Name3'",
                )
                pool["exclude_if_day"] = _str_to_exclude_if_day(eid_str)

                # Pool removal button - only shown on last pool if there's more than 1
                if i == len(pools) - 1 and len(pools) > 1:
                    if st.button(
                        "Pool entfernen", 
                        key=f"{prefix}_p{i}_remove",
                        help="Entfernt den letzten Pool.",
                    ):
                        pools.pop(i)
                        cfg["pools"] = pools
                        save_meeting_pools(pools_data)
                        st.rerun()

            if st.button("Pool hinzufügen", key=f"{prefix}_add_pool"):
                exclude_laufen_global = st.session_state.get("global_exclude_laufen", False)
                pools.append({
                    "type": "names",
                    "names": [],
                    "site": cfg.get("site", "BH"),
                    "exclude_laufen": exclude_laufen_global,
                })
                cfg["pools"] = pools
                save_meeting_pools(pools_data)
                st.rerun()

            # Fallback fields
            st.markdown("---")
            st.markdown("**Fallback-Einstellungen**")
            c1, c2 = st.columns(2)
            cfg["fallback_text"] = c1.text_input(
                "Fallback-Text", value=cfg.get("fallback_text", "FÄLLT AUS"),
                key=f"{prefix}_fb_text",
            )
            cfg["roter_fallback_text"] = c2.checkbox(
                "Roter Text", value=cfg.get("roter_fallback_text", True),
                key=f"{prefix}_fb_rot",
            )

            cfg["pools"] = pools

    st.divider()
    if st.button("Alle Pool-Änderungen speichern", type="primary", key="save_pools_btn"):
        # Validate pools before saving
        validation_errors = []
        for meeting_key, cfg in pools_data.items():
            for i, pool in enumerate(cfg.get("pools", []), start=1):
                if pool.get("type") == "names" and not pool.get("names"):
                    validation_errors.append(f"{meeting_key} - Pool {i}: Typ 'Person' hat keine Namen ausgewählt. Bitte Namen auswählen oder Typ ändern.")
                if pool.get("type") == "group" and not pool.get("group"):
                    validation_errors.append(f"{meeting_key} - Pool {i}: Typ 'Gruppe' hat keine Gruppe ausgewählt.")
        
        if validation_errors:
            st.error("**Validierungsfehler - Speichern nicht möglich:**")
            for err in validation_errors:
                st.error(f"• {err}")
        else:
            # Clean up None/empty values before saving
            for meeting_key, cfg in pools_data.items():
                for pool in cfg.get("pools", []):
                    for k in list(pool.keys()):
                        if pool[k] is None or pool[k] == "" or pool[k] == [] or pool[k] is False:
                            if k not in ("type", "names", "group", "site", "exclude_laufen"):
                                del pool[k]
            save_meeting_pools(pools_data)
            st.success("Rapporte-Pools gespeichert und neu geladen.")
            st.rerun()


# ===========================================================================
# TAB 6 — Radiologe in Laufen
# ===========================================================================

with tab_laufen:
    st.subheader("Radiologe in Laufen")
    st.caption("Konfiguration für den Standort Laufen.")

    laufen_days = st.multiselect(
        "Radiologe in Laufen anwesend",
        options=ALL_WEEKDAYS,
        default=list(sched.LAUFEN_DAYS),
        help="Wochentage, an denen 'Laufen' besetzt wird (OG-Leader Neuro/Laufen).",
        key="laufen_days_select"
    )
    st.session_state["laufen_days"] = laufen_days

    # Derive current global exclude_laufen from first pool of first rapport
    _pools_data_laufen = load_meeting_pools()
    _first_cfg = next(iter(_pools_data_laufen.values()), {})
    _first_pool = (_first_cfg.get("pools") or [{}])[0]
    _current_excl = _first_pool.get("exclude_laufen", False)

    global_exclude_laufen = st.checkbox(
        "Laufen von allen Rapporte-Pools ausschließen",
        value=st.session_state.get("global_exclude_laufen", _current_excl),
        help="Schließt den Radiologen in Laufen von der Zuteilung in allen Rapport-Pools aus.",
        key="global_exclude_laufen",
    )

    if st.button("Laufen-Einstellungen speichern", type="primary", key="save_laufen_btn"):
        # Update LAUFEN_DAYS
        sched.LAUFEN_DAYS.clear()
        sched.LAUFEN_DAYS.update(laufen_days)
        # Apply exclude_laufen to all pools in all rapporte
        _pools_data_laufen = load_meeting_pools()
        for _cfg in _pools_data_laufen.values():
            for _pool in _cfg.get("pools", []):
                _pool["exclude_laufen"] = global_exclude_laufen
        save_meeting_pools(_pools_data_laufen)
        st.success("Laufen-Einstellungen gespeichert und auf alle Pools angewendet.")
        st.rerun()


