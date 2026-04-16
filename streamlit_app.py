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
    for s in sorted(sched.staff_by_name.values(), key=lambda x: (x.role, x.name)):
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
        pw = st.text_input("Passwort", type="password", label_visibility="collapsed",
                           placeholder="Passwort eingeben")
        if st.button("Anmelden", use_container_width=True, type="primary"):
            if pw == st.secrets.get("password", ""):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Falsches Passwort.")
    return False

if not _check_password():
    st.stop()



st.set_page_config(
    page_title="Wochenplan Scheduler",
    page_icon="📋",
    layout="wide",
)
st.title("Wochenplan Scheduler — KSBL Radiologie")

tab_plan, tab_personal, tab_layout = st.tabs(["Wochenplan", "Personalverwaltung", "Layout-Editor"])

# ===========================================================================
# TAB 1 — Wochenplan
# ===========================================================================

with tab_plan:
    uploaded_file = st.file_uploader(
        "Leeren .xlsm-Wochenplan hochladen",
        type=["xlsm"],
        help="Die Wochenplan-Vorlage mit bereits eingetragenen Absenzen und Diensten.",
    )

    col1, col2 = st.columns([1, 2])
    with col1:
        seed = st.number_input("Seed", value=1234, step=1, format="%d",
                               help="Zufalls-Seed für reproduzierbare Ergebnisse.")
    with col2:
        laufen_days = st.multiselect(
            "Radiologe in Laufen anwesend",
            options=ALL_WEEKDAYS,
            default=["Dienstag"],
        )

    run_btn = st.button(
        "Pipeline starten",
        disabled=(uploaded_file is None),
        type="primary",
    )

    if run_btn and uploaded_file is not None:
        input_tmp_path = None
        output_tmp_path = None
        try:
            # Write upload to a named temp file (openpyxl needs a path)
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f_in:
                f_in.write(uploaded_file.getbuffer())
                input_tmp_path = f_in.name

            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f_out:
                output_tmp_path = f_out.name

            with st.spinner("Pipeline läuft…"):
                # Configure Laufen days for this run
                sched.LAUFEN_DAYS.clear()
                sched.LAUFEN_DAYS.update(laufen_days)

                # Reset all counters for a clean run
                sched.reset_all_counters()

                # Load workbook
                wb = load_workbook(input_tmp_path, data_only=False, keep_vba=True)
                ws = wb["Wochenplan"]

                # Stage 0: cleanup
                sched.cleanup_blocks(ws, clear_fr=True, clear_og=True, clear_meetings=True)

                # Stage 1: read absences
                absences = sched.read_absences_by_day(ws)

                # Stage 2: OG leaders
                sched.assign_la_to_ogs(ws, absences)

                # Stage 3: OG non-leaders + coverage flags
                sched.assign_nonleaders_to_ogs(ws, absences, seed=int(seed))

                # Stage 4: FR shifts
                sched.assign_fr_shifts_to_cells(ws, absences, seed=int(seed))

                # Stage 5: meetings
                sched.assign_meetings(ws, absences, seed=int(seed))

                # Save and restore dropped OOXML parts
                wb.save(output_tmp_path)
                sched.patch_xlsm(output_tmp_path, input_tmp_path)

                with open(output_tmp_path, "rb") as f:
                    result_bytes = f.read()

            base = uploaded_file.name.rsplit(".", 1)[0]
            st.session_state["result_bytes"] = result_bytes
            st.session_state["result_filename"] = f"{base}_FINAL.xlsm"
            st.success("Pipeline erfolgreich abgeschlossen.")

        except KeyError as e:
            st.error(
                f"Blatt 'Wochenplan' nicht gefunden oder unerwartete Struktur. "
                f"Details: {e}"
            )
            st.session_state["result_bytes"] = None
        except Exception as e:
            st.error(f"Fehler während der Pipeline: {e}")
            st.session_state["result_bytes"] = None
        finally:
            for p in (input_tmp_path, output_tmp_path):
                if p and os.path.exists(p):
                    try:
                        os.unlink(p)
                    except OSError:
                        pass

    # Download button and stats — persisted in session_state across re-runs
    if st.session_state["result_bytes"] is not None:
        st.download_button(
            label="Ergebnis herunterladen (.xlsm)",
            data=st.session_state["result_bytes"],
            file_name=st.session_state["result_filename"],
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )

        st.subheader("Wochenstatistik")
        stats_rows = [
            {
                "Name": s.name,
                "Rolle": s.role,
                "Standort": s.site,
                "Frontarzt": s.fr_shifts_count,
                "Rapporte": s.meetings_count,
            }
            for s in sorted(sched.staff_by_name.values(), key=lambda x: x.name)
            if s.fr_shifts_count > 0 or s.meetings_count > 0
        ]
        if stats_rows:
            st.dataframe(
                pd.DataFrame(stats_rows),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Alle Zählerstände sind 0 — Pipeline wurde noch nicht ausgeführt.")

# ===========================================================================
# TAB 2 — Personalverwaltung
# ===========================================================================

def _staff_form(form_key: str, defaults: dict | None = None) -> dict | None:
    """Render a staff edit/add form and return the submitted values as a dict,
    or None if the form was not submitted."""
    d = defaults or {}
    with st.form(form_key, clear_on_submit=(defaults is None)):
        c1, c2, c3 = st.columns(3)
        name     = c1.text_input("Name", value=d.get("name", ""), placeholder="J. Beispiel",
                                 disabled=(defaults is not None))
        role     = c2.selectbox("Rolle", ["AA", "FA", "LA"],
                                index=["AA", "FA", "LA"].index(d.get("role", "AA")))
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
        fr_always = fr_col1.checkbox(
            "Nie Frontarzt",
            value=d.get("fr_excluded", False),
        )
        fr_days = fr_col2.multiselect(
            "Nur an diesen Tagen kein Frontarzt",
            options=sched.WEEKDAYS,
            default=sorted(d.get("fr_excluded_days", [])),
            disabled=fr_always,
            help="Wird ignoriert wenn 'Nie Frontarzt' aktiviert ist.",
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


with tab_personal:
    st.subheader("Personalbestand")

    # Read-only overview table
    st.dataframe(
        staff_to_display_dataframe(),
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
            del sched.staff_by_name[edit_name]
            sched.rebuild_quick_views()
            save_staff_to_json()
            st.success(f"'{edit_name}' wurde entfernt.")
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
# TAB 3 — Layout-Editor
# ===========================================================================

_COL_CFG = {
    day: st.column_config.TextColumn(day[:2], help=day)
    for day in ALL_WEEKDAYS
}

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

    # ---- Save ----
    st.divider()
    if st.button("Alle Änderungen speichern", type="primary", key="save_layout_btn"):
        new_layout = {
            "abw_ranges": {
                row: abw_edited.at[row, "Bereich"] for row in abw_edited.index
            },
            "nacht_ranges": {
                row: nacht_edited.at[row, "Bereich"] for row in nacht_edited.index
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
        }
        save_layout(new_layout)
        st.success("Layout gespeichert und neu geladen.")
        st.rerun()
