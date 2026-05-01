# -*- coding: utf-8 -*-
"""
Wochenplan Scheduler (pipeline: OG leaders → OG non-leaders → FR → Meetings)

- OG leaders (LA) → assigned to ALL of their dedicated OGs when present
  (includes Laufen; H.W. Ott leads Neuro & Laufen; active days defined by layout.json cells)

- OG non-leaders (FAs & AAs) → rotations first, then balance; coverage flags:
    • WENIGER ALS 2FA for MSK/Neuro/Onko/Thorax/Abdomen if total FAs < 2
    • KEIN FA IN BH / KEIN FA IN LI only for MSK & Abdomen (suppressed if < 2 FA flagged)
    • KEIN AA for all OGs except those in og_rules.json (rotation_or_leader_only)

- FR (Frontarzt) AFTER OGs
    • FR-only absences overlay includes that day’s Laufen assignees
    • optional manual FR-only exclusions (e.g., {"Donnerstag": {"H.W. Ott"}})

- Meetings LAST; standardized 4-pool priority matcher; excludes Laufen assignees per-day;
  Medizin BH & LI on Monday colored red (text only); per-meeting rules as configured.

Keeps macros/formatting (keep_vba=True). Uses exact cell maps.
"""

from dataclasses import dataclass, field
from typing import Set, Dict, List, Tuple, Optional
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
import random, re, os, zipfile, json, xml.etree.ElementTree as ET
from pathlib import Path

# ---------------- Constants ----------------

BH = "BH"
LI = "LI"

AA = "AA"   # Assistenzarzt
OA = "OA"   # Oberarzt
LA = "LA"   # Leitender Arzt

WEEKDAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]

# OG special rules — loaded from og_rules.json and organgruppen.json (edit via Streamlit UI)
def _load_og_rules():
    _path = Path(__file__).parent / "og_rules.json"
    with open(_path, encoding="utf-8") as _f:
        _r = json.load(_f)
    
    # Load OG list from organgruppen.json
    _og_path = Path(__file__).parent / "organgruppen.json"
    global OG_LIST_NO_LAUFEN, OG_LIST
    if _og_path.exists():
        with open(_og_path, encoding="utf-8") as _og_f:
            _og_data = json.load(_og_f)
            OG_LIST_NO_LAUFEN = _og_data.get("organgruppen", [])
            OG_LIST = OG_LIST_NO_LAUFEN  # No special Laufen handling
    else:
        # Fallback to hardcoded list if file doesn't exist
        OG_LIST_NO_LAUFEN = ["MSK", "Neuro", "Onko", "Thorax", "Abdomen", "Mammo", "Intervention/ Vaskulär", "Nuklearmedizin", "Laufen"]
        OG_LIST = OG_LIST_NO_LAUFEN
    
    global OG_PRIORITY_ORDER, USE_RANDOM_OG_SELECTION, OG_WEIGHTS_OA, OG_WEIGHTS_AA, OG_MAX_FAS, OG_MAX_AAS
    OG_PRIORITY_ORDER = _r.get("og_priority_order", OG_LIST_NO_LAUFEN)
    USE_RANDOM_OG_SELECTION = _r.get("use_random_og_selection", False)
    
    # Load separate OA and AA weights
    OG_WEIGHTS_OA = _r.get("og_weights_oa", _r.get("og_weights", {}))  # Fallback to old og_weights if needed
    OG_WEIGHTS_AA = _r.get("og_weights_aa", _r.get("og_weights", {}))
    
    OG_MAX_FAS = _r.get("og_max_fas", {})
    OG_MAX_AAS = _r.get("og_max_aas", {})
    
    # Set defaults for any missing OGs
    for og in OG_LIST:
        if og not in OG_WEIGHTS_OA:
            OG_WEIGHTS_OA[og] = 0.4 if og in ["Mammo", "Intervention/ Vaskulär"] else 0.6
        if og not in OG_WEIGHTS_AA:
            OG_WEIGHTS_AA[og] = 0.4 if og in ["Mammo", "Intervention/ Vaskulär"] else 0.6
        if og not in OG_MAX_FAS:
            OG_MAX_FAS[og] = None  # No limit
        if og not in OG_MAX_AAS:
            OG_MAX_AAS[og] = None  # No limit
    
    return (
        set(_r.get("rotation_or_leader_only", [])),
        set(_r.get("warn_kein_aa", [])),
        set(_r.get("warn_weniger_als_2fa", [])),
        set(_r.get("warn_kein_fa_site", [])),
    )

# OG Lists - will be populated from organgruppen.json
OG_LIST_NO_LAUFEN: List[str] = []
OG_LIST: List[str] = []

OG_PRIORITY_ORDER: List[str] = []
USE_RANDOM_OG_SELECTION: bool = False
OG_WEIGHTS_OA: Dict[str, float] = {}
OG_WEIGHTS_AA: Dict[str, float] = {}
OG_MAX_FAS: Dict[str, Optional[int]] = {}
OG_MAX_AAS: Dict[str, Optional[int]] = {}
OG_ROTATION_OR_LEADER_ONLY, OG_WARN_KEIN_AA, TARGET_OG_FOR_ONE_FA, TARGET_OG_FOR_KEIN_FA_SITE = _load_og_rules()
OGS_SKIP_KEIN_AA = set(OG_LIST) - OG_WARN_KEIN_AA



# ---------------- Data model ----------------

@dataclass
class Staff:
    name: str
    role: str            # AA | OA | LA
    site: str            # BH | LI
    leads_ogs: Set[str] = field(default_factory=set)        # for LA only
    rotations: Set[str] = field(default_factory=set)        # AA + FA(non-leader)
    fr_excluded: bool = False                               # if True → never Frontarzt on any day
    fr_excluded_days: Set[str] = field(default_factory=set) # excluded on specific weekdays only
    is_cover: bool = False                                  # if True → Stellvertreter (not shown in absence list)

    # Counters
    meetings_count: int = 0  # Deprecated - kept for backwards compatibility
    meetings_count_week: int = 0
    meetings_count_montag: int = 0
    meetings_count_dienstag: int = 0
    meetings_count_mittwoch: int = 0
    meetings_count_donnerstag: int = 0
    meetings_count_freitag: int = 0
    fr_shifts_count: int = 0
    og_nonleader_count: int = 0
    aa_og_count: int = 0

    @property
    def is_fa(self) -> bool:
        return self.role in {OA, LA}

staff_by_name: Dict[str, Staff] = {}

def add_staff(name: str, role: str, site: str,
              leads_for: Optional[List[str]] = None,
              rotation: Optional[List[str]] = None,
              fr_excluded: bool = False,
              fr_excluded_days: Optional[List[str]] = None,
              is_cover: bool = False):
    leads = set(leads_for or []) if role == LA else set()
    rots  = set(rotation or [])
    staff_by_name[name] = Staff(
        name=name,
        role=role,
        site=site,
        leads_ogs=leads,
        rotations=rots,
        fr_excluded=fr_excluded,
        fr_excluded_days=set(fr_excluded_days or []),
        is_cover=is_cover,
    )

def rebuild_quick_views() -> None:
    """Rebuild the module-level quick-view lists from staff_by_name.
    Must be called after any mutation of staff_by_name."""
    global aa_bh, aa_li, oa_bh, oa_li
    global fa_all_bh, fa_all_li, la_bh, la_li, leaders_by_og
    aa_bh            = [s.name for s in staff_by_name.values() if s.role == AA and s.site == BH]
    aa_li            = [s.name for s in staff_by_name.values() if s.role == AA and s.site == LI]
    oa_bh            = [s.name for s in staff_by_name.values() if s.role == OA and s.site == BH]
    oa_li            = [s.name for s in staff_by_name.values() if s.role == OA and s.site == LI]
    fa_all_bh        = [s.name for s in staff_by_name.values() if s.is_fa and s.site == BH]
    fa_all_li        = [s.name for s in staff_by_name.values() if s.is_fa and s.site == LI]
    la_bh            = [s.name for s in staff_by_name.values() if s.role == LA and s.site == BH]
    la_li            = [s.name for s in staff_by_name.values() if s.role == LA and s.site == LI]
    leaders_by_og    = {
        og: sorted([s.name for s in staff_by_name.values() if og in s.leads_ogs])
        for og in OG_LIST
    }


def load_staff_from_json(path: str) -> None:
    """Load staff from a JSON file, replacing the current staff_by_name contents.
    JSON format: list of objects with keys name, role, site, leads_ogs, rotations, fr_excluded, is_cover.
    Also loads site_rules from the same JSON file.
    Rebuilds all quick-view lists automatically."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    
    # Load site_rules if present
    global SITE_RULES
    if isinstance(data, dict) and "site_rules" in data:
        SITE_RULES = data.get("site_rules", {})
        records = data.get("staff", [])
    else:
        # Old format: just array of staff
        SITE_RULES = {"BH": {"no_oa_vormittag": False}, "LI": {"no_oa_vormittag": True}}
        records = data
    
    staff_by_name.clear()
    for r in records:
        add_staff(
            name=r["name"],
            role=r["role"],
            site=r["site"],
            leads_for=r.get("leads_ogs", []),
            rotation=r.get("rotations", []),
            fr_excluded=r.get("fr_excluded", False),
            fr_excluded_days=r.get("fr_excluded_days", []),
            is_cover=r.get("is_cover", False),
        )
    rebuild_quick_views()


# Quick-view placeholders — populated by load_staff_from_json below
aa_bh = aa_li = oa_bh = oa_li = []
fa_all_bh = fa_all_li = la_bh = la_li = []
leaders_by_og: Dict[str, List[str]] = {}

# Site rules loaded from staff.json
SITE_RULES: Dict[str, Dict[str, bool]] = {}

# staff.json is the single source of truth — always load from it.
_staff_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "staff.json")
if not os.path.exists(_staff_json):
    raise FileNotFoundError(
        f"staff.json not found at {_staff_json}. "
        "Create it before running the scheduler."
    )
load_staff_from_json(_staff_json)

# ---------------- Sheet utils & ranges ----------------

# Populated by load_layout_from_json() at startup — do not edit here.
ABW_RANGES:        Dict[str, str]                                      = {}
NACHT_RANGES:      Dict[str, str]                                      = {}
SPAETDIENST_CELLS: Dict[str, Dict[str, str]]                           = {}
FR_CELLS:          Dict[str, Dict[str, Tuple[str, ...]]]               = {}
OG_CELLS:          Dict[str, Dict[str, Tuple[str, ...]]]               = {}
MEETING_CELLS:     Dict[str, Dict[str, Dict[str, Tuple[str, ...]]]]    = {}
MEDIZIN_MONDAY_CELLS: Dict[str, str]                                   = {}
VORDERGRUNDDIENST_CELLS: Dict[str, str]                                = {}
HINTERGRUNDDIENST_CELLS: Dict[str, str]                                = {}
DATE_CELLS:        Dict[str, str]                                      = {}
WEEKDAY_DATE_CELLS: Dict[str, str]                                     = {}
FEIERTAGE:         Set[str]                                            = set()
FEIERTAGE_MERGE_CELLS: Dict[str, str]                                  = {}


def load_layout_from_json(path: str) -> None:
    """Load all cell-map constants from layout.json, replacing current values."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    ABW_RANGES.clear()
    ABW_RANGES.update(data["abw_ranges"])

    NACHT_RANGES.clear()
    NACHT_RANGES.update(data["nacht_ranges"])

    SPAETDIENST_CELLS.clear()
    for site, day_map in data["spaetdienst_cells"].items():
        SPAETDIENST_CELLS[site] = dict(day_map)

    FR_CELLS.clear()
    for site, day_map in data["fr_cells"].items():
        FR_CELLS[site] = {day: tuple(cells) for day, cells in day_map.items()}

    OG_CELLS.clear()
    for og, day_map in data["og_cells"].items():
        OG_CELLS[og] = {day: tuple(cells) for day, cells in day_map.items()}

    MEETING_CELLS.clear()
    for site, mtg_map in data["meeting_cells"].items():
        MEETING_CELLS[site] = {
            mtg: {day: tuple(cells) for day, cells in day_map.items()}
            for mtg, day_map in mtg_map.items()
        }

    MEDIZIN_MONDAY_CELLS.clear()
    MEDIZIN_MONDAY_CELLS.update(data.get("medizin_monday_cells", {}))

    VORDERGRUNDDIENST_CELLS.clear()
    VORDERGRUNDDIENST_CELLS.update(data.get("vordergrunddienst_cells", {}))

    HINTERGRUNDDIENST_CELLS.clear()
    HINTERGRUNDDIENST_CELLS.update(data.get("hintergrunddienst_cells", {}))

    DATE_CELLS.clear()
    DATE_CELLS.update(data.get("date_cells", {}))

    WEEKDAY_DATE_CELLS.clear()
    WEEKDAY_DATE_CELLS.update(data.get("weekday_date_cells", {}))

    FEIERTAGE.clear()
    FEIERTAGE.update(data.get("feiertage", []))

    FEIERTAGE_MERGE_CELLS.clear()
    FEIERTAGE_MERGE_CELLS.update(data.get("feiertage_merge_cells", {}))


# layout.json holds all Excel cell-coordinate maps.
_layout_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "layout.json")
if not os.path.exists(_layout_json):
    raise FileNotFoundError(
        f"layout.json not found at {_layout_json}. "
        "Create it before running the scheduler."
    )
load_layout_from_json(_layout_json)


# meeting_pools.json holds the priority-pool definitions for every Rapport.
MEETING_POOLS: Dict[str, dict] = {}

def load_meeting_pools_from_json(path: str) -> None:
    """Load meeting pool definitions from JSON, replacing current MEETING_POOLS."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    MEETING_POOLS.clear()
    MEETING_POOLS.update(data)

_pools_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "meeting_pools.json")
if not os.path.exists(_pools_json):
    raise FileNotFoundError(
        f"meeting_pools.json not found at {_pools_json}. "
        "Create it before running the scheduler."
    )
load_meeting_pools_from_json(_pools_json)


def reload_og_rules() -> None:
    """Reload OG special rules from og_rules.json and organgruppen.json — call after saving via UI."""
    global OG_ROTATION_OR_LEADER_ONLY, OG_WARN_KEIN_AA, TARGET_OG_FOR_ONE_FA, TARGET_OG_FOR_KEIN_FA_SITE, OGS_SKIP_KEIN_AA
    OG_ROTATION_OR_LEADER_ONLY, OG_WARN_KEIN_AA, TARGET_OG_FOR_ONE_FA, TARGET_OG_FOR_KEIN_FA_SITE = _load_og_rules()
    OGS_SKIP_KEIN_AA = set(OG_LIST) - OG_WARN_KEIN_AA


def _clear_cells(ws: Worksheet, cells: Tuple[str, ...]):
    for a1 in cells:
        ws[a1].value = ""

def cleanup_blocks(ws: Worksheet, *, clear_fr=True, clear_og=True, clear_meetings=True):
    """
    Clears the configured cells for FR, OG, and/or Meetings.
    Formats/macros are preserved (caller is responsible for saving the workbook).
    """
    if clear_fr:
        for site in FR_CELLS:
            for day in WEEKDAYS:
                _clear_cells(ws, FR_CELLS[site][day])

    if clear_og:
        for og in OG_CELLS:
            for day in WEEKDAYS:
                _clear_cells(ws, OG_CELLS[og][day])

    if clear_meetings:
        for site in MEETING_CELLS:
            for mtg in MEETING_CELLS[site]:
                for day in WEEKDAYS:
                    _clear_cells(ws, MEETING_CELLS[site][mtg].get(day, ()))
    
def tokens_from_val(val: str) -> List[str]:
    if not isinstance(val, str) or not val.strip(): return []
    return [p.strip() for p in re.split(r",|;|/|\n|•", val) if p.strip()]

def read_spaetdienst_by_day(ws: Worksheet) -> Dict[str, Dict[str, Set[str]]]:
    """
    Returns {"BH": {weekday: set(names)}, "LI": {weekday: set(names)}} for Spätdienst,
    reading from fixed cells defined in SPAETDIENST_CELLS.
    """
    out: Dict[str, Dict[str, Set[str]]] = {
        "BH": {d: set() for d in WEEKDAYS},
        "LI": {d: set() for d in WEEKDAYS},
    }

    for site, day_map in SPAETDIENST_CELLS.items():
        for day, a1 in day_map.items():
            v = ws[a1].value
            if v:
                for t in tokens_from_val(str(v)):
                    out[site][day].add(t)

    return out

def _add_from_range(ws: Worksheet, a1: str, dest: set):
    for row in ws[a1]:
        for cell in row:
            if cell.value:
                for t in tokens_from_val(str(cell.value)):
                    dest.add(t)

def read_absences_by_day(ws: Worksheet) -> Dict[str, Set[str]]:
    """Read absences from ABW_RANGES and NACHT_RANGES cells."""
    absences = {d: set() for d in WEEKDAYS}
    for d in WEEKDAYS:
        _add_from_range(ws, ABW_RANGES[d], absences[d])
        _add_from_range(ws, NACHT_RANGES[d], absences[d])
    return absences

def remove_covers_from_absences_visual(ws: Worksheet) -> None:
    """
    Remove covers (Stellvertreter) from the visual absence list in ABW_RANGES.
    This is called AFTER the pipeline has finished, purely for display purposes.
    The pipeline itself uses the full absence list (including covers).
    """
    # Find all staff who are covers
    cover_names = {s.name for s in staff_by_name.values() if s.is_cover}
    
    if not cover_names:
        return  # Nothing to do
    
    # For each day's absence range
    for day in WEEKDAYS:
        if day not in ABW_RANGES:
            continue
        
        cell_range = ABW_RANGES[day]
        
        # Read current names from the range
        current_names = []
        for row in ws[cell_range]:
            for cell in row:
                val = cell.value
                if val and isinstance(val, str):
                    for name in tokens_from_val(val):
                        if name:
                            current_names.append(name)
        
        # Filter out covers
        filtered_names = [n for n in current_names if n not in cover_names]
        
        # Clear ALL cells in range
        for row in ws[cell_range]:
            for cell in row:
                cell.value = None
        
        # Write back filtered list (from top)
        cells_list = [cell for row in ws[cell_range] for cell in row]
        for i, name in enumerate(filtered_names):
            if i < len(cells_list):
                cells_list[i].value = name

def _make_font(old: Font, *, bold: bool, color: str) -> Font:
    """
    Return a new Font that copies every attribute from *old* but overrides
    bold and color.  color must be a fully-specified 8-char ARGB hex string
    (e.g. "FFFF0000" for opaque red, "FF000000" for opaque black) so that
    openpyxl does not silently prepend "00" and produce a transparent colour.
    """
    return Font(
        name=old.name, size=old.size, bold=bold, italic=old.italic,
        underline=old.underline, strike=old.strike, vertAlign=old.vertAlign,
        charset=old.charset, scheme=old.scheme, family=old.family,
        outline=old.outline, shadow=old.shadow, condense=old.condense,
        extend=old.extend, color=color,
    )

def set_bold_red(ws: Worksheet, a1: str):
    ws[a1].font = _make_font(ws[a1].font or Font(), bold=True,  color="FFFF0000")

def set_black_normal(ws: Worksheet, a1: str):
    ws[a1].font = _make_font(ws[a1].font or Font(), bold=False, color="FF000000")

def set_red(ws: Worksheet, a1: str):
    ws[a1].font = _make_font(ws[a1].font or Font(), bold=False, color="FFFF0000")
    
def reset_all_counters():
    """
    Reset all per-person and per-meeting counters so that each run of the
    pipeline starts from a clean state.
    """
    # Per-person counters
    for s in staff_by_name.values():
        s.meetings_count = 0
        s.meetings_count_week = 0
        s.meetings_count_montag = 0
        s.meetings_count_dienstag = 0
        s.meetings_count_mittwoch = 0
        s.meetings_count_donnerstag = 0
        s.meetings_count_freitag = 0
        s.fr_shifts_count = 0
        s.og_nonleader_count = 0
        s.aa_og_count = 0

    # Per-meeting/pool counters
    POOL_COUNTS.clear()

    
def print_weekly_stats():
    """
    Print a simple overview of Frontarztdienste (FR) and Rapporte (meetings)
    per person for this run.
    """
    print("\n=== Wochenstatistik: Frontarztdienste & Rapporte ===")
    print(f"{'Name':25} {'Frontarzt':>3} {'Rapporte':>9}")
    print("-" * 40)
    for s in sorted(staff_by_name.values(), key=lambda x: x.name):
        # You can show everyone, or only those with at least 1 of either:
        if s.fr_shifts_count == 0 and s.meetings_count == 0:
            continue
        print(f"{s.name:25} {s.fr_shifts_count:>3} {s.meetings_count:>9}")
    print("-" * 40)

    

# ---------------- FR (Frontarzt) ----------------

def get_persons_assigned_to_laufen(ws: Worksheet) -> Dict[str, Set[str]]:
    """
    Read names listed in the Laufen OG cells.
    Returns: {weekday -> set(names)}
    """
    out: Dict[str, Set[str]] = {d: set() for d in WEEKDAYS}

    cells_map = OG_CELLS.get("Laufen", {})
    for day in WEEKDAYS:
        cells = cells_map.get(day, tuple())
        # Defensive: coerce single strings to tuples if someone edits accidentally
        if isinstance(cells, str):
            cells = (cells,)
        for a1 in cells:
            if not isinstance(a1, str):
                continue
            v = ws[a1].value
            if isinstance(v, str):
                name = v.strip()
                if name:
                    out[day].add(name)
    return out

def absences_for_fr_stage(ws: Worksheet, absences_orig: Dict[str, Set[str]],
                          extra_exclusions: Optional[Dict[str, Set[str]]] = None,
                          include_laufen_from_og: bool = True) -> Dict[str, Set[str]]:
    abs_fr = {d: set(absences_orig.get(d, set())) for d in WEEKDAYS}

    if include_laufen_from_og:
        laufen = get_persons_assigned_to_laufen(ws)
        for d in WEEKDAYS:
            abs_fr[d].update(laufen.get(d, set()))

    if extra_exclusions:
        for d in WEEKDAYS:
            abs_fr[d].update(extra_exclusions.get(d, set()))

    return abs_fr

def pick_fa_for_fr_shift(day: str, fa_pool: list,
                          absences_by_day: dict,
                          avoid=None,
                          rng: random.Random = random):
    avoid = set(avoid or [])
    present = []
    for n in fa_pool:
        if n in avoid:
            continue
        if n in absences_by_day.get(day, set()):
            continue
        s = staff_by_name.get(n)
        if not s:
            continue
        if s.fr_excluded or day in s.fr_excluded_days:
            continue
        present.append(n)

    if not present:
        return None, "no eligible FR candidate"

    minc = min(staff_by_name[n].fr_shifts_count for n in present)
    bucket = [n for n in present if staff_by_name[n].fr_shifts_count == minc]
    choice = rng.choice(bucket)
    staff_by_name[choice].fr_shifts_count += 1
    return choice, None

def assign_fr_shifts_to_cells(
    ws: Worksheet,
    absences_orig: Dict[str, Set[str]],
    rng: random.Random,
    *,
    extra_exclusions: Optional[Dict[str, Set[str]]] = None,
    include_laufen_from_og: bool = True,
) -> None:
    abs_fr = absences_for_fr_stage(ws, absences_orig,
                                   extra_exclusions=extra_exclusions,
                                   include_laufen_from_og=include_laufen_from_og)

    for day in WEEKDAYS:
        if day in FEIERTAGE:
            continue  # Skip holidays
        
        # BH: Process all cells
        used = set()
        bh_no_oa_vormittag = SITE_RULES.get("BH", {}).get("no_oa_vormittag", False)
        
        if bh_no_oa_vormittag and FR_CELLS["BH"][day]:
            # First cell: LA only
            top_cell = FR_CELLS["BH"][day][0]
            pick, _ = pick_fa_for_fr_shift(day, la_bh, abs_fr, used, rng)
            ws[top_cell].value = pick or ""
            if pick:
                used.add(pick)
            # Remaining cells: all FA
            for a1 in FR_CELLS["BH"][day][1:]:
                pick, _ = pick_fa_for_fr_shift(day, fa_all_bh, abs_fr, used, rng)
                ws[a1].value = pick or ""
                if pick:
                    used.add(pick)
        else:
            # All cells: all FA
            for a1 in FR_CELLS["BH"][day]:
                pick, _ = pick_fa_for_fr_shift(day, fa_all_bh, abs_fr, used, rng)
                ws[a1].value = pick or ""
                if pick:
                    used.add(pick)

        # LI: Same logic
        used = set()
        li_no_oa_vormittag = SITE_RULES.get("LI", {}).get("no_oa_vormittag", True)
        
        if li_no_oa_vormittag and FR_CELLS["LI"][day]:
            # First cell: LA only
            top_cell = FR_CELLS["LI"][day][0]
            pick, _ = pick_fa_for_fr_shift(day, la_li, abs_fr, used, rng)
            ws[top_cell].value = pick or ""
            if pick:
                used.add(pick)
            # Remaining cells: all FA
            for a1 in FR_CELLS["LI"][day][1:]:
                pick, _ = pick_fa_for_fr_shift(day, fa_all_li, abs_fr, used, rng)
                ws[a1].value = pick or ""
                if pick:
                    used.add(pick)
        else:
            # All cells: all FA
            for a1 in FR_CELLS["LI"][day]:
                pick, _ = pick_fa_for_fr_shift(day, fa_all_li, abs_fr, used, rng)
                ws[a1].value = pick or ""
                if pick:
                    used.add(pick)



# ---------------- OG assignment ----------------

FA_COUNTS: Dict[str, Dict[str,int]] = {d:{og:0 for og in OG_LIST} for d in WEEKDAYS}
AA_COUNTS: Dict[str, Dict[str,int]] = {d:{og:0 for og in OG_LIST} for d in WEEKDAYS}

def reset_og_counts():
    for d in WEEKDAYS:
        for og in OG_LIST: 
            FA_COUNTS[d][og]=0
            AA_COUNTS[d][og]=0

def _first_empty_cell(ws: Worksheet, cells: Tuple[str,...]) -> Optional[str]:
    for a1 in cells:
        v = ws[a1].value
        if v is None or (isinstance(v,str) and v.strip()==""): return a1
    return None

def _already_listed(ws: Worksheet, cells: Tuple[str,...], name: str) -> bool:
    for a1 in cells:
        v = ws[a1].value
        if isinstance(v,str) and v.strip()==name: return True
    return False

def assign_la_to_ogs(ws: Worksheet, absences_by_day: Dict[str, Set[str]]) -> Dict[str, Dict[str, int]]:
    """
    Assign every present LA to all of their dedicated OGs (no prioritization).
    OGs without cells for a day are skipped automatically.
    """
    reset_og_counts()

    for day in WEEKDAYS:
        if day in FEIERTAGE:
            continue  # Skip holidays
        abs_today = absences_by_day.get(day, set())

        for og in OG_LIST:
            cells = OG_CELLS[og][day]
            if not cells:
                continue

            for leader in leaders_by_og.get(og, []):
                if leader in abs_today:
                    continue
                if _already_listed(ws, cells, leader):
                    continue
                slot = _first_empty_cell(ws, cells)
                if slot:
                    ws[slot].value = leader
                    FA_COUNTS[day][og] += 1

    return FA_COUNTS

def _names_in_cells(ws: Worksheet, cells: Tuple[str,...]) -> List[str]:
    out=[]
    for a1 in cells:
        v = ws[a1].value
        if isinstance(v,str) and v.strip(): out.append(v.strip())
    return out

def _has_fa_from_site(ws: Worksheet, cells: Tuple[str,...], site: str) -> bool:
    for nm in _names_in_cells(ws,cells):
        s = staff_by_name.get(nm)
        if s and s.is_fa and s.site==site: return True
    return False

def _has_aa(ws: Worksheet, cells: Tuple[str,...]) -> bool:
    for nm in _names_in_cells(ws,cells):
        s = staff_by_name.get(nm)
        if s and s.role==AA: return True
    return False

def _place_in_og(ws: Worksheet, day: str, og: str, name: str, count_for_fa: bool):
    cells = OG_CELLS[og][day]
    if not cells: return False
    if _already_listed(ws,cells,name): return False
    slot = _first_empty_cell(ws,cells)
    if not slot: return False
    ws[slot].value = name
    s = staff_by_name.get(name)
    # Note: og_nonleader_count and aa_og_count are now incremented with weights
    # in assign_nonleaders_to_ogs, not here
    if count_for_fa and s and s.role==OA:
        FA_COUNTS[day][og] += 1
    if not count_for_fa and s and s.role==AA:
        AA_COUNTS[day][og] += 1
    return True

def assign_nonleaders_to_ogs(ws: Worksheet, absences_by_day: Dict[str,Set[str]], rng: random.Random) -> Dict[str,Dict[str,int]]:
    """
    Assign non-leader FAs and AAs to OGs using hybrid OG-centric approach with weights.
    
    Strategy:
    1. Reset daily counters for each person
    2. For OAs: While pool not empty, pick OG with lowest FA_COUNT, assign best matching OA
    3. For AAs: Same process with separate counter
    4. Respects OG_WEIGHTS (e.g., 0.4 for Mammo allows 2 assignments, 0.6 allows 1)
    """
    for day in WEEKDAYS:
        if day in FEIERTAGE:
            continue  # Skip holidays
        abs_today = absences_by_day.get(day, set())
        
        # ===== ROUND 1: OAs =====
        present_oas = [n for n in (oa_bh + oa_li) if n not in abs_today]
        
        # Reset daily counters
        for name in present_oas:
            staff_by_name[name].og_nonleader_count = 0
        
        # Pool: All OAs with counter that can still fit smallest weight
        min_weight = min(OG_WEIGHTS_OA.values()) if OG_WEIGHTS_OA else 0.6
        pool = set(present_oas)
        
        while pool:
            # 1. Find OGs with free slots and under max_fas limit
            available_ogs = [og for og in OG_LIST 
                           if og not in OG_ROTATION_OR_LEADER_ONLY
                           and _first_empty_cell(ws, OG_CELLS[og][day]) is not None
                           and (OG_MAX_FAS.get(og) is None or FA_COUNTS[day][og] < OG_MAX_FAS[og])]
            
            if not available_ogs:
                break  # No free slots
            
            # 2. For each OG: Find compatible persons (counter + weight <= 1.0 AND not already in this OG)
            og_candidates = {}
            for og in available_ogs:
                og_weight = OG_WEIGHTS_OA.get(og, 0.6)
                compatible = [n for n in pool 
                            if staff_by_name[n].og_nonleader_count + og_weight <= 1.0
                            and not _already_listed(ws, OG_CELLS[og][day], n)]
                if compatible:
                    og_candidates[og] = compatible
            
            if not og_candidates:
                break  # No compatible OG-person combinations
            
            # 3. Prioritize OGs that have people with rotations waiting
            # First, try to find OGs where someone has a rotation match
            ogs_with_rotation_matches = []
            for og, candidates in og_candidates.items():
                if any(og in staff_by_name[n].rotations for n in candidates):
                    ogs_with_rotation_matches.append(og)
            
            if ogs_with_rotation_matches:
                # Choose among OGs with rotation matches, preferring lowest FA_COUNT
                eligible_ogs = ogs_with_rotation_matches
            else:
                # No rotation matches available, use all OGs
                eligible_ogs = list(og_candidates.keys())
            
            # 4. Choose OG with lowest FA_COUNT (among eligible)
            minv = min(FA_COUNTS[day][og] for og in eligible_ogs)
            bucket = [og for og in eligible_ogs if FA_COUNTS[day][og] == minv]
            
            # 5. OG-Priority or Random
            if USE_RANDOM_OG_SELECTION:
                chosen_og = rng.choice(bucket)
            else:
                chosen_og = sorted(bucket, key=lambda x: OG_PRIORITY_ORDER.index(x) if x in OG_PRIORITY_ORDER else 999)[0]
            
            # 6. Choose person from compatible candidates
            # Prioritize: in_rotation > no_rotation > other_rotation
            candidates = og_candidates[chosen_og]
            
            in_rotation = [n for n in candidates if chosen_og in staff_by_name[n].rotations]
            no_rotation = [n for n in candidates if not staff_by_name[n].rotations]
            other_rotation = [n for n in candidates 
                            if staff_by_name[n].rotations 
                            and chosen_og not in staff_by_name[n].rotations]
            
            if in_rotation:
                pick = rng.choice(in_rotation)
            elif no_rotation:
                pick = rng.choice(no_rotation)
            elif other_rotation:
                pick = rng.choice(other_rotation)
            else:
                break  # Should not happen
            
            # 7. Place person in OG
            _place_in_og(ws, day, chosen_og, pick, count_for_fa=True)
            
            # 8. Update counter
            og_weight = OG_WEIGHTS_OA.get(chosen_og, 0.6)
            staff_by_name[pick].og_nonleader_count += og_weight
            
            # 9. Remove from pool if can't fit any more OGs
            if staff_by_name[pick].og_nonleader_count + min_weight > 1.0:
                pool.discard(pick)
        
        # ===== ROUND 2: AAs =====
        present_aas = [n for n in (aa_bh + aa_li) if n not in abs_today]
        
        # Reset daily counters
        for name in present_aas:
            staff_by_name[name].aa_og_count = 0
        
        pool = set(present_aas)
        
        while pool:
            # Same logic as OAs, but using aa_og_count and checking max_aas
            available_ogs = [og for og in OG_LIST 
                           if og not in OG_ROTATION_OR_LEADER_ONLY
                           and _first_empty_cell(ws, OG_CELLS[og][day]) is not None
                           and (OG_MAX_AAS.get(og) is None or AA_COUNTS[day][og] < OG_MAX_AAS[og])]
            
            if not available_ogs:
                break
            
            og_candidates = {}
            for og in available_ogs:
                og_weight = OG_WEIGHTS_AA.get(og, 0.6)
                compatible = [n for n in pool 
                            if staff_by_name[n].aa_og_count + og_weight <= 1.0
                            and not _already_listed(ws, OG_CELLS[og][day], n)]
                if compatible:
                    og_candidates[og] = compatible
            
            if not og_candidates:
                break
            
            # Prioritize OGs with rotation matches
            ogs_with_rotation_matches = []
            for og, candidates in og_candidates.items():
                if any(og in staff_by_name[n].rotations for n in candidates):
                    ogs_with_rotation_matches.append(og)
            
            if ogs_with_rotation_matches:
                eligible_ogs = ogs_with_rotation_matches
            else:
                eligible_ogs = list(og_candidates.keys())
            
            minv = min(AA_COUNTS[day][og] for og in eligible_ogs)
            bucket = [og for og in eligible_ogs if AA_COUNTS[day][og] == minv]
            
            if USE_RANDOM_OG_SELECTION:
                chosen_og = rng.choice(bucket)
            else:
                chosen_og = sorted(bucket, key=lambda x: OG_PRIORITY_ORDER.index(x) if x in OG_PRIORITY_ORDER else 999)[0]
            
            candidates = og_candidates[chosen_og]
            
            in_rotation = [n for n in candidates if chosen_og in staff_by_name[n].rotations]
            no_rotation = [n for n in candidates if not staff_by_name[n].rotations]
            other_rotation = [n for n in candidates 
                            if staff_by_name[n].rotations 
                            and chosen_og not in staff_by_name[n].rotations]
            
            if in_rotation:
                pick = rng.choice(in_rotation)
            elif no_rotation:
                pick = rng.choice(no_rotation)
            elif other_rotation:
                pick = rng.choice(other_rotation)
            else:
                break
            
            _place_in_og(ws, day, chosen_og, pick, count_for_fa=False)
            
            og_weight = OG_WEIGHTS_AA.get(chosen_og, 0.6)
            staff_by_name[pick].aa_og_count += og_weight
            
            if staff_by_name[pick].aa_og_count + min_weight > 1.0:
                pool.discard(pick)
        
        # ===== ROUND 3: Coverage flags =====
        for og in OG_LIST:
            cells = OG_CELLS[og][day]
            wrote_fa_shortage = False
            if og in TARGET_OG_FOR_ONE_FA and FA_COUNTS[day][og] < 2:
                slot = _first_empty_cell(ws,cells)
                if slot: ws[slot].value = "WENIGER ALS 2FA"; set_bold_red(ws,slot)
                wrote_fa_shortage = True
            if not wrote_fa_shortage and og in TARGET_OG_FOR_KEIN_FA_SITE:
                if not _has_fa_from_site(ws,cells,BH):
                    slot=_first_empty_cell(ws,cells)
                    if slot: ws[slot].value="KEIN FA IN BH"; set_bold_red(ws,slot)
                if not _has_fa_from_site(ws,cells,LI):
                    slot=_first_empty_cell(ws,cells)
                    if slot: ws[slot].value="KEIN FA IN LI"; set_bold_red(ws,slot)
            if og not in OGS_SKIP_KEIN_AA and not _has_aa(ws,cells):
                slot=_first_empty_cell(ws,cells)
                if slot: ws[slot].value="KEIN AA"; set_bold_red(ws,slot)
    
    return FA_COUNTS

# ---------------- Meetings (Rapporte): standardized 4-pool system ----------------

# Fairness counters for meeting pools
POOL_COUNTS = defaultdict(int)  # key: (meeting_key, pool_index, name) -> count

def _group_names(group: str, site: str) -> List[str]:
    if group == "AA":
        return aa_bh if site == BH else aa_li
    if group == "OA":
        return oa_bh if site == BH else oa_li
    if group == "LA":
        return [n for n in staff_by_name if staff_by_name[n].role == LA and staff_by_name[n].site == site]
    if group == "FA_ALL":
        return fa_all_bh if site == BH else fa_all_li
    raise ValueError(f"Unknown group: {group}")

def _filter_candidates(base: List[str], *, day: str, site: str,
                       absences: Dict[str, Set[str]],
                       spaetdienst_by_site_day: Dict[str, Dict[str, Set[str]]],
                       exclude_spaetdienst: Optional[str] = None,
                       exclude_names: Optional[Set[str]] = None,
                       exclude_if_day: Optional[Dict[str, Set[str]]] = None,
                       exclude_laufen: bool = False,
                       laufen_names: Optional[Set[str]] = None) -> List[str]:
    c = [n for n in base if n not in absences.get(day, set())]
    if exclude_spaetdienst:
        c = [n for n in c if n not in spaetdienst_by_site_day.get(exclude_spaetdienst, {}).get(day, set())]
    if exclude_names:
        c = [n for n in c if n not in exclude_names]
    if exclude_if_day and day in exclude_if_day:
        c = [n for n in c if n not in exclude_if_day[day]]
    if exclude_laufen and laufen_names:
        c = [n for n in c if n not in laufen_names]
    return c

def _fair_pick_pool(meeting_key: str, pool_idx: int, candidates: List[str], 
                    rng: random.Random, day: str) -> Optional[str]:
    """
    Fair pick for a given meeting pool with per-day and per-week balancing:
      1) Prefer people with FEWER meetings TODAY (meetings_count_<day>)
      2) Among those, prefer FEWER meetings THIS WEEK (meetings_count_week)
      3) Random choice among equals (using SEED for reproducibility)
    
    Note: POOL_COUNTS is no longer used - day and week counters provide better fairness.
    """
    if not candidates:
        return None

    # Step 1: Prefer people with fewer meetings today
    day_counter_name = f"meetings_count_{day.lower()}"
    min_today = min(getattr(staff_by_name[n], day_counter_name, 0) for n in candidates)
    cand_by_today = [n for n in candidates if getattr(staff_by_name[n], day_counter_name, 0) == min_today]

    # Step 2: Among those, prefer people with fewer meetings this week
    min_week = min(staff_by_name[n].meetings_count_week for n in cand_by_today)
    bucket = [n for n in cand_by_today if staff_by_name[n].meetings_count_week == min_week]

    # Step 3: Random choice with SEED
    return rng.choice(bucket)

def _bump_pool(meeting_key: str, pool_idx: int, name: str):
    POOL_COUNTS[(meeting_key, pool_idx, name)] += 1

def _assign(ws: Worksheet, a1: str, text: str, style: Optional[str]=None):
    ws[a1].value = text
    if style=="red_bold": set_bold_red(ws, a1)
    elif style=="black": set_black_normal(ws, a1)

def write_medizin_placeholders_monday(ws: Worksheet):
    for a1 in MEDIZIN_MONDAY_CELLS.values():
        if a1 and isinstance(a1, str) and a1.strip():
            _assign(ws, a1, "BITTE EINTRAGEN", "red_bold")
    
def assign_meeting_by_pools(
    ws: Worksheet,
    *,
    rng: random.Random,
    meeting_key: str,          # e.g. "BH|Medizin (07:45-08:00)"
    site: str,                 # BH or LI
    day: str,
    cells: Tuple[str, ...],
    pools: List[dict],         # up to 4 pools
    absences: Dict[str, Set[str]],
    spaetdienst: Dict[str, Dict[str, Set[str]]],
    laufen_names: Set[str],
    monday_style: Optional[str] = None,       # e.g., "red"
    fallback_text: Optional[str] = "FÄLLT AUS",
    fallback_style: Optional[str] = "red_bold",
):
    for a1 in cells:
        placed = False
        for idx, pool in enumerate(pools, start=1):
            ptype = pool.get("type")
            style = pool.get("style")

            # Expand candidates for this pool
            if ptype == "names":
                base = list(pool.get("names", []))
            elif ptype == "group":
                base = _group_names(pool["group"], pool.get("site", site))
            elif ptype == "spaetdienst_aa":
                base = list(spaetdienst[pool.get("site", site)][day])
                base = [n for n in base if staff_by_name.get(n) and staff_by_name[n].role == AA and staff_by_name[n].site == pool.get("site", site)]
            else:
                raise ValueError(f"Unknown pool type: {ptype}")

            # Filter (supports per-day Laufen exclusion)
            cands = _filter_candidates(
                base, day=day, site=site, absences=absences,
                spaetdienst_by_site_day=spaetdienst,
                exclude_spaetdienst=pool.get("exclude_spaetdienst"),
                exclude_names=set(pool.get("exclude_names", [])) or None,
                exclude_if_day={k:set(v) for k,v in pool.get("exclude_if_day", {}).items()} if pool.get("exclude_if_day") else None,
                exclude_laufen=pool.get("exclude_laufen", False),
                laufen_names=laufen_names,
            )

            pick = _fair_pick_pool(meeting_key, idx, cands, rng, day)
            if pick:
                _assign(ws, a1, pick, style)
                if monday_style and day == "Montag":
                    if monday_style == "red": set_red(ws, a1)
                _bump_pool(meeting_key, idx, pick)
                
                # Increment counters
                staff_by_name[pick].meetings_count += 1  # Deprecated but kept for stats
                staff_by_name[pick].meetings_count_week += 1
                day_counter_name = f"meetings_count_{day.lower()}"
                current_count = getattr(staff_by_name[pick], day_counter_name, 0)
                setattr(staff_by_name[pick], day_counter_name, current_count + 1)
                
                placed = True
                break

        if not placed and fallback_text is not None:
            _assign(ws, a1, fallback_text, fallback_style)

def assign_meetings(ws: Worksheet, absences: Dict[str, Set[str]], rng: random.Random):
    write_medizin_placeholders_monday(ws)

    laufen_by_day = get_persons_assigned_to_laufen(ws)
    spaet         = read_spaetdienst_by_day(ws)

    # Iterate over all meetings defined in meeting_pools.json
    for meeting_key, cfg in MEETING_POOLS.items():
        site = cfg["site"]
        # Derive the meeting name from the key (format: "SITE|Meeting Name")
        mtg_name = meeting_key.split("|", 1)[1]

        # Look up the cell map in MEETING_CELLS
        mtg_cells = MEETING_CELLS.get(site, {}).get(mtg_name, {})
        if not mtg_cells:
            continue

        pools              = cfg.get("pools", [])
        fallback_text      = cfg.get("fallback_text", "FÄLLT AUS")
        roter_fallback     = cfg.get("roter_fallback_text", True)
        fallback_style     = "red_bold" if roter_fallback else "black"

        for day, cells in mtg_cells.items():
            if day in FEIERTAGE:
                continue  # Skip holidays
            if not cells:
                continue

            assign_meeting_by_pools(
                ws, rng=rng, meeting_key=meeting_key, site=site, day=day, cells=cells,
                pools=pools, absences=absences, spaetdienst=spaet,
                laufen_names=laufen_by_day.get(day, set()),
                monday_style=None,
                fallback_text=fallback_text, fallback_style=fallback_style,
            )



# ---------------- XLSM post-save patch ----------------

def patch_xlsm(output_path: str, input_path: str) -> None:
    """
    Restore parts that openpyxl silently drops when saving a .xlsm file.

    openpyxl rewrites xl/worksheets/sheet1.xml and xl/styles.xml correctly,
    but it also:
      • Drops xl/drawings/drawing2.xml and its rels (strips the chart for sheet2)
      • Rewrites xl/worksheets/_rels/sheet2.xml.rels to point at drawing1.xml
        instead of drawing2.xml (wrong file, chart disappears)
      • Drops xl/printerSettings/printerSettings1.bin and its rels reference
      • Drops xl/sharedStrings.xml and xl/calcChain.xml
      • Changes [Content_Types].xml Default for .bin from printerSettings to vbaProject
      • Changes the VML relationship ID in sheet1.xml from "rId3" to "anysvml"

    Strategy: rebuild the output archive starting from the INPUT (which has the
    correct structure), then override only the parts openpyxl legitimately changed:
    sheet1.xml (cell data) and styles.xml (font/style definitions). Rebuild
    sheet1.xml.rels to include both the input's non-VML relationships and openpyxl's
    "anysvml" VML ID (which is what the new sheet1.xml references).
    """
    VML_TYPE = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
    )
    RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    def read_archive(path: str) -> dict:
        data = {}
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                data[name] = zf.read(name)
        return data

    inp = read_archive(input_path)
    out = read_archive(output_path)

    # Start from input (correct rels, drawings, Content_Types, printerSettings, etc.)
    patched = dict(inp)

    # Override the parts openpyxl correctly rewrites
    for key in ("xl/worksheets/sheet1.xml", "xl/styles.xml"):
        if key in out:
            patched[key] = out[key]

    # Rebuild sheet1.xml.rels: keep all input relationships except VML,
    # then add VML with the "anysvml" ID that openpyxl embedded in sheet1.xml.
    inp_s1_rels = ET.fromstring(inp["xl/worksheets/_rels/sheet1.xml.rels"])
    rel_parts = []
    for rel in inp_s1_rels:
        if rel.get("Type") != VML_TYPE:
            rel_parts.append(
                f'<Relationship Id="{rel.get("Id")}"'
                f' Type="{rel.get("Type")}"'
                f' Target="{rel.get("Target")}"/>'
            )
    rel_parts.append(
        f'<Relationship Id="anysvml"'
        f' Type="{VML_TYPE}"'
        f' Target="../drawings/vmlDrawing1.vml"/>'
    )
    merged_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{RELS_NS}">'
        + "".join(rel_parts)
        + "</Relationships>"
    )
    patched["xl/worksheets/_rels/sheet1.xml.rels"] = merged_rels.encode("utf-8")

    # Write the patched archive back to output_path
    tmp = output_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in patched.items():
            zout.writestr(name, data)
    os.replace(tmp, output_path)


# ---------------- CLI pipeline (OG → OG non-leaders → FR → Meetings) ----------------
# Default: bash> python wochenplan_scheduler.py -i "KW XX_leer.xlsm" -o "KW XX.xlsm" --seed 1234
# Skip cleanup: append -nocleanup
# keep intermediates: append --keep

# ===========================================================================
# CSV Import Functions
# ===========================================================================

def match_csv_name_to_staff(csv_name: str) -> Optional[str]:
    """
    Match a CSV name to the exact staff.json name by surname, with initial fallback for duplicates.
    
    CSV format: "Nachname(n) Vorname(n)" - last part is always first name
    staff.json format: "Initial(s) Nachname" - last part is always surname
    
    Strategy:
    1. Take all parts from CSV except the last (those are surname parts)
    2. Split by hyphens to get all surname candidates
    3. Match against staff.json surnames (last part of staff name)
    4. If multiple matches, filter by first initial
    5. Return exact staff.json name
    
    Examples:
        "Iacobut Andreea-Emilia" → "A. Iacobut"
        "Slezenkovska-Stojanoska Tanja" → "T. Slezenkovska"
        "Berberich Bettina Katharina" → "B. Berberich"
        "Ott Hans-Werner" → "H.W. Ott"
    """
    csv_parts = csv_name.strip().split()
    if len(csv_parts) < 2:
        return None
    
    # All parts except last = surname parts
    surname_parts = csv_parts[:-1]
    # Last part = first name
    first_name = csv_parts[-1]
    
    # Get first initial from first name (handle hyphens)
    first_initial = first_name.split('-')[0][0].upper() if first_name else None
    
    # Split surnames by hyphens and collect all surname candidates
    surname_candidates = []
    for part in surname_parts:
        if '-' in part:
            surname_candidates.extend(part.split('-'))
        else:
            surname_candidates.append(part)
    
    # Normalize: lowercase for comparison
    surname_candidates_lower = [s.lower() for s in surname_candidates]
    
    # Find all staff with matching surname
    matches = []
    for staff_name in staff_by_name.keys():
        staff_parts = staff_name.split()
        if len(staff_parts) < 2:
            continue
        
        # Last part of staff name is the surname
        staff_surname = staff_parts[-1].lower()
        
        if staff_surname in surname_candidates_lower:
            matches.append(staff_name)
    
    # If single match, return it
    if len(matches) == 1:
        return matches[0]
    
    # If multiple matches, filter by first initial
    if len(matches) > 1 and first_initial:
        initial_matches = [m for m in matches if m.split()[0][0].upper() == first_initial]
        if len(initial_matches) == 1:
            return initial_matches[0]
        elif len(initial_matches) > 1:
            # Still ambiguous - return first and warn
            print(f"Warning: Ambiguous match for {csv_name}: {initial_matches}, using {initial_matches[0]}")
            return initial_matches[0]
    
    # No match or still ambiguous
    if matches:
        print(f"Warning: Multiple surname matches for {csv_name}: {matches}, using {matches[0]}")
        return matches[0]
    
    print(f"Warning: No staff match found for CSV name: {csv_name}")
    return None


def fill_dienste_from_csv(ws: Worksheet, csv_path: str) -> None:
    """
    Read dienste and absences from CSV and fill into Excel template.
    
    Args:
        ws: openpyxl worksheet to fill
        csv_path: Path to CSV file with dienste data
    """
    import pandas as pd
    from datetime import datetime
    
    # Try UTF-8 first (better special character support), fallback to ISO-8859-1
    try:
        df = pd.read_csv(csv_path, sep=';', encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv(csv_path, sep=';', encoding='ISO-8859-1')
    
    # Parse dates
    df['Datum'] = pd.to_datetime(df['Datum'], format='%d.%m.%Y')
    
    # Find Monday of the week (first day)
    mondays = df[df['Datum'].dt.dayofweek == 0]['Datum']
    if len(mondays) == 0:
        raise ValueError("No Monday found in CSV - cannot determine week start")
    monday_date = mondays.iloc[0]
    
    # Write date to T20 (if date_cells exist in layout)
    if DATE_CELLS and "first_monday" in DATE_CELLS:
        ws[DATE_CELLS["first_monday"]].value = monday_date.strftime('%d/%m/%Y')
    
    # Calculate and write KW (ISO week number)
    if DATE_CELLS and "kw_number" in DATE_CELLS:
        kw = monday_date.isocalendar()[1]
        ws[DATE_CELLS["kw_number"]].value = f"KW {kw}"
    
    # Dienst type mapping — loaded from bezeichnungen.json
    _bez_path = Path(__file__).parent / "bezeichnungen.json"
    with open(_bez_path, encoding="utf-8") as _f:
        _bez = json.load(_f)
    ABSENZ_TYPES = set(_bez.get("absenz", []))
    
    # Day name mapping (German date to layout key)
    day_names = {
        0: "Montag", 1: "Dienstag", 2: "Mittwoch",
        3: "Donnerstag", 4: "Freitag", 5: "Samstag", 6: "Sonntag"
    }
    
    # Collect data by day and type
    absences_by_day = {d: [] for d in day_names.values()}
    nacht_by_day = {}
    spaet_by_day = {"BH": {}, "LI": {}}
    vordergrund_by_day = {}
    hintergrund_by_day = {}
    
    # Process each row
    for _, row in df.iterrows():
        dienst_type = row['Bezeichnung']
        csv_name = row['Suchname']
        date = row['Datum']
        day_name = day_names[date.dayofweek]
        
        # Convert name format - match to staff.json
        matched_name = match_csv_name_to_staff(csv_name)
        if not matched_name:
            print(f"Skipping unknown person from CSV: {csv_name}")
            continue
        abbrev_name = matched_name
        
        # Categorize dienst
        if dienst_type in ABSENZ_TYPES:
            absences_by_day[day_name].append(abbrev_name)
        
        elif "Nachtdienst" in dienst_type:
            nacht_by_day[day_name] = abbrev_name
        
        elif "Spätdienst" in dienst_type:
            site = "BH" if dienst_type.startswith("Bh-") else "LI"
            spaet_by_day[site][day_name] = abbrev_name
        
        # Vordergrunddienst - specific weekend dienste
        elif "Pikett_Vormittag_Sa/So" in dienst_type:
            vordergrund_by_day[day_name] = abbrev_name
        
        elif "Tagdienst Sa/So" in dienst_type:
            vordergrund_by_day[day_name] = abbrev_name
        
        # Hintergrunddienst - all Pikett types (including Pikett_24h_Sa/So)
        elif "Pikett" in dienst_type:
            hintergrund_by_day[day_name] = abbrev_name
    
    # Write to Excel
    from openpyxl.styles import PatternFill
    gray_fill = PatternFill(start_color="D0CECE", end_color="D0CECE", fill_type="solid")
    
    # Absences - write each name to separate row (skip on holidays)
    for day, names in absences_by_day.items():
        if day in FEIERTAGE:
            continue  # Skip absences on holidays
        if names and day in ABW_RANGES:
            cell_range = ABW_RANGES[day]
            # Parse range: "T94:AC108" → start_col="T", start_row=94
            from openpyxl.utils import column_index_from_string, get_column_letter
            
            start_cell = cell_range.split(':')[0]
            # Extract column letter and row number
            col_letter = ''.join(c for c in start_cell if c.isalpha())
            start_row = int(''.join(c for c in start_cell if c.isdigit()))
            
            # Write each name to a new row
            for idx, name in enumerate(sorted(set(names))):
                cell = f"{col_letter}{start_row + idx}"
                ws[cell].value = name
    
    # Nachtdienst - single cell per day (write on all days including holidays)
    for day, name in nacht_by_day.items():
        if day in NACHT_RANGES:
            cell_range = NACHT_RANGES[day]
            first_cell = cell_range.split(':')[0]
            ws[first_cell].value = name
    
    # Spätdienst (normal days)
    for site in ["BH", "LI"]:
        for day, name in spaet_by_day[site].items():
            if day in FEIERTAGE:
                continue  # Handled separately below
            cell = SPAETDIENST_CELLS.get(site, {}).get(day)
            if cell:
                ws[cell].value = name
    
    # Vordergrunddienst (holidays and weekends)
    for day, name in vordergrund_by_day.items():
        if day in FEIERTAGE:
            # On holidays: Merge cells using FEIERTAGE_MERGE_CELLS
            merge_range = FEIERTAGE_MERGE_CELLS.get(day)
            if merge_range:
                ws.merge_cells(merge_range)
                # Write to first cell of range
                first_cell = merge_range.split(':')[0]
                ws[first_cell].value = name
        else:
            # Normal weekend: Write to Vordergrunddienst cell
            cell = VORDERGRUNDDIENST_CELLS.get(day)
            if cell:
                ws[cell].value = name
    
    # Hintergrunddienst (write on all days including holidays)
    for day, name in hintergrund_by_day.items():
        cell = HINTERGRUNDDIENST_CELLS.get(day)
        if cell:
            ws[cell].value = name
    
    # Gray out FR, OG, Rapporte, and Absences on holidays
    for day in FEIERTAGE:
        # Gray out Absence range
        if day in ABW_RANGES:
            cell_range = ABW_RANGES[day]
            # Parse range to get all cells
            start_cell, end_cell = cell_range.split(':')
            
            # Extract column letters and row numbers
            start_col = ''.join(c for c in start_cell if c.isalpha())
            start_row = int(''.join(c for c in start_cell if c.isdigit()))
            end_col = ''.join(c for c in end_cell if c.isalpha())
            end_row = int(''.join(c for c in end_cell if c.isdigit()))
            
            # Gray out all cells in the range
            from openpyxl.utils import column_index_from_string
            start_col_idx = column_index_from_string(start_col)
            end_col_idx = column_index_from_string(end_col)
            
            for col_idx in range(start_col_idx, end_col_idx + 1):
                for row in range(start_row, end_row + 1):
                    from openpyxl.utils import get_column_letter
                    cell = f"{get_column_letter(col_idx)}{row}"
                    ws[cell].fill = gray_fill
        
        # Gray out FR cells
        for site in ["BH", "LI"]:
            fr_cells = FR_CELLS.get(site, {}).get(day, [])
            for cell in fr_cells:
                ws[cell].fill = gray_fill
        
        # Gray out OG cells
        for og in OG_CELLS:
            og_cells = OG_CELLS[og].get(day, [])
            for cell in og_cells:
                ws[cell].fill = gray_fill
        
        # Gray out Rapporte cells
        for site in ["BH", "LI"]:
            for mtg_name, mtg_days in MEETING_CELLS.get(site, {}).items():
                mtg_cells = mtg_days.get(day, [])
                for cell in mtg_cells:
                    ws[cell].fill = gray_fill


if __name__ == "__main__":
    import argparse, os

    parser = argparse.ArgumentParser(
        description="Run Wochenplan pipeline: OGs (LA→non-LA/AA) → FR → Meetings."
    )
    parser.add_argument("-i", "--input",  required=True,
                        help=".xlsm source file (e.g. KW 41_leer.xlsm)")
    parser.add_argument("-o", "--output", required=True,
                        help="FINAL .xlsm (after meetings), e.g. KW_41_FINAL.xlsm")
    parser.add_argument("--csv", dest="csv_file", default=None,
                        help="CSV file with dienste to import (optional)")
    parser.add_argument("--seed", type=int, default=1234,
                        help="random seed for fair picks (default: 1234)")
    parser.add_argument("-nocleanup", "--no-cleanup", dest="no_cleanup",
                        action="store_true",
                        help="Skip pre-run cleanup of FR/OG/Meetings cells (default: cleanup runs)")
    parser.add_argument("--keep-intermediate", "--keep", dest="keep_intermediate",
                        action="store_true",
                        help="Keep intermediate .xlsm files (default: delete them)")

    args = parser.parse_args()

    # Start with fresh counters
    reset_all_counters()
    
    # Create single RNG for entire pipeline (reproducible with seed)
    rng = random.Random(args.seed)

    # Intermediate filenames (only used when --keep-intermediate is set)
    out_stem       = os.path.splitext(args.output)[0]
    cleaned_out    = f"{out_stem}_0_CLEANED.xlsm"
    og_leaders_out = f"{out_stem}_1_OG_LEADERS.xlsm"
    og_full_out    = f"{out_stem}_2_OG_FULL.xlsm"
    fr_out         = f"{out_stem}_3_FR_APPLIED.xlsm"

    # Load workbook once
    wb = load_workbook(args.input, data_only=False, keep_vba=True)
    ws = wb["Wochenplan"]

    # 0) Cleanup (default)
    if args.no_cleanup:
        cleaned_note = "SKIPPED"
    else:
        cleanup_blocks(ws, clear_fr=True, clear_og=True, clear_meetings=True)
        cleaned_note = cleaned_out
        if args.keep_intermediate:
            wb.save(cleaned_out)
            patch_xlsm(cleaned_out, args.input)

    # 0.5) Import dienste from CSV (if provided)
    if args.csv_file:
        print(f"📥 Importing dienste from CSV: {args.csv_file}")
        fill_dienste_from_csv(ws, args.csv_file)
        print("✓ CSV import complete")

    # Read absences once (cleanup does not touch absence cells)
    absences = read_absences_by_day(ws)

    # 1) OG leaders (LA)
    assign_la_to_ogs(ws, absences)
    if args.keep_intermediate:
        wb.save(og_leaders_out)
        patch_xlsm(og_leaders_out, args.input)

    # 2) OG non-leaders (FA/AA) + flags
    assign_nonleaders_to_ogs(ws, absences, rng)
    if args.keep_intermediate:
        wb.save(og_full_out)
        patch_xlsm(og_full_out, args.input)

    # 3) FR (depends on OG + Laufen in OG)
    assign_fr_shifts_to_cells(ws, absences, rng)
    if args.keep_intermediate:
        wb.save(fr_out)
        patch_xlsm(fr_out, args.input)

    # 4) Meetings
    assign_meetings(ws, absences, rng)
    
    # 5) Remove covers from visual absence list (purely cosmetic)
    remove_covers_from_absences_visual(ws)

    # Save final output once and patch
    wb.save(args.output)
    patch_xlsm(args.output, args.input)

    print("✔ Pipeline completed")
    print(f"  0_CLEANUP      → {cleaned_note}")
    if args.keep_intermediate:
        print(f"  1_OG_LA        → {og_leaders_out}")
        print(f"  2_OG_FULL      → {og_full_out}")
        print(f"  3_FR_APPLIED   → {fr_out}")
    print(f"  4_+MEETINGS    → {args.output}")

    # Print per-person statistics (FR + Rapporte)
    print_weekly_stats()

    if args.keep_intermediate:
        print("↪ Kept intermediates (user requested).")
