# -*- coding: utf-8 -*-
"""
Wochenplan Scheduler (pipeline: OG leaders → OG non-leaders → FR → Meetings)

- OG leaders (LA) → assigned to ALL of their dedicated OGs when present
  (includes Laufen; H.W. Ott leads Neuro & Laufen; Laufen active days = LAUFEN_DAYS)

- OG non-leaders (FAs & AAs) → rotations first, then balance; coverage flags:
    • WENIGER ALS 2FA for MSK/Neuro/Onko/Thorax/Abdomen if total FAs < 2
    • KEIN FA IN BH / KEIN FA IN LI only for MSK & Abdomen (suppressed if < 2 FA flagged)
    • KEIN AA for all OGs except Nuklearmedizin & Laufen

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

# ---------------- Constants ----------------

BH = "BH"
LI = "LI"

AA = "AA"   # Assistenzarzt
OA = "OA"   # Oberarzt
LA = "LA"   # Leitender Arzt

# Laufen enabled weekdays (tweakable)
LAUFEN_DAYS: Set[str] = {"Dienstag"}

OG_LIST = [
    "MSK","Neuro","Onko","Thorax","Abdomen","Mammo","Intervention/ Vaskulär","Nuklearmedizin","Laufen"
]
WEEKDAYS = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]

# Flags for OGs
TARGET_OG_FOR_ONE_FA = {"MSK","Neuro","Onko","Thorax","Abdomen"}  # < 2 FA total → WENIGER ALS 2FA
TARGET_OG_FOR_KEIN_FA_SITE = {"MSK","Abdomen"}                    # KEIN FA IN BH/LI (only if not <2FA flagged)
OGS_SKIP_KEIN_AA = {"Nuklearmedizin","Laufen", "Mammo"}           # no KEIN AA flag

# Non-leader OGs (exclude Laufen)
OG_LIST_NONLEADER = [og for og in OG_LIST if og != "Laufen"]



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
    absent_by_default: bool = False                         # if True → absent unless covers_for person is absent
    covers_for: Optional[str] = None                        # name of the person this staff member stands in for

    # Counters
    meetings_count: int = 0
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
              absent_by_default: bool = False,
              covers_for: Optional[str] = None):
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
        absent_by_default=absent_by_default,
        covers_for=covers_for,
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
    JSON format: list of objects with keys name, role, site, leads_ogs, rotations, fr_excluded.
    Rebuilds all quick-view lists automatically."""
    with open(path, encoding="utf-8") as f:
        records = json.load(f)
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
            absent_by_default=r.get("absent_by_default", False),
            covers_for=r.get("covers_for", None),
        )
    rebuild_quick_views()


# Quick-view placeholders — populated by load_staff_from_json below
aa_bh = aa_li = oa_bh = oa_li = []
fa_all_bh = fa_all_li = la_bh = la_li = []
leaders_by_og: Dict[str, List[str]] = {}

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

def apply_substitution_rules(
    absences_by_day: Dict[str, Set[str]]
) -> Dict[str, Set[str]]:
    """
    Apply absent_by_default / covers_for rules from staff_by_name.

    For each staff member with absent_by_default=True:
      - They start absent every day.
      - If their covers_for person is absent that day, they become available
        (removed from the absence set), UNLESS they are also explicitly listed
        as absent in the Excel file.
    """
    # Collect names explicitly absent in the Excel (before any rule is applied)
    explicit: Dict[str, Set[str]] = {
        day: set(names) for day, names in absences_by_day.items()
    }

    adjusted: Dict[str, Set[str]] = {
        day: set(names) for day, names in absences_by_day.items()
    }

    for s in staff_by_name.values():
        if not s.absent_by_default:
            continue
        for day in WEEKDAYS:
            day_set = adjusted.setdefault(day, set())
            explicit_day = explicit.get(day, set())

            if s.covers_for and s.covers_for in day_set:
                # The person they cover for is absent → they may work
                # Only make them available if they are not explicitly absent
                if s.name not in explicit_day:
                    day_set.discard(s.name)
            else:
                # Covered person is present → stand-in stays absent
                day_set.add(s.name)

    return adjusted


def read_absences_by_day(ws: Worksheet) -> Dict[str, Set[str]]:
    absences = {d: set() for d in WEEKDAYS}
    for d in WEEKDAYS:
        _add_from_range(ws, ABW_RANGES[d], absences[d])
        _add_from_range(ws, NACHT_RANGES[d], absences[d])
    return apply_substitution_rules(absences)

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

def read_laufen_by_day(ws: Worksheet, allowed_days: Optional[Set[str]] = None) -> Dict[str, Set[str]]:
    """
    Read names listed in the Laufen OG cells.
    If allowed_days is provided, only those weekdays are considered.
    Returns: {weekday -> set(names)}
    """
    out: Dict[str, Set[str]] = {d: set() for d in WEEKDAYS}
    days = WEEKDAYS if allowed_days is None else [d for d in WEEKDAYS if d in allowed_days]

    cells_map = OG_CELLS.get("Laufen", {})
    for day in days:
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
        laufen = read_laufen_by_day(ws, allowed_days=LAUFEN_DAYS)
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
    seed: Optional[int] = None,
    *,
    extra_exclusions: Optional[Dict[str, Set[str]]] = None,
    include_laufen_from_og: bool = True,
) -> None:
    rng = random.Random(seed)
    abs_fr = absences_for_fr_stage(ws, absences_orig,
                                   extra_exclusions=extra_exclusions,
                                   include_laufen_from_og=include_laufen_from_og)

    for day in WEEKDAYS:
        # BH: unchanged
        used = set()
        for a1 in FR_CELLS["BH"][day]:
            pick, _ = pick_fa_for_fr_shift(day, fa_all_bh, abs_fr, used, rng)
            ws[a1].value = pick or ""
            if pick:
                used.add(pick)

        # LI: top line = LA only, bottom line = FA/LA as before
        used = set()
        cells_li = FR_CELLS["LI"][day]
        if cells_li:
            # First cell (row 32) → only LAs (la_li)
            top_cell = cells_li[0]
            pick, _ = pick_fa_for_fr_shift(day, la_li, abs_fr, used, rng)
            ws[top_cell].value = pick or ""
            if pick:
                used.add(pick)

            # Remaining cells (row 33 etc.) → any FR-eligible FA/LA in LI
            for a1 in cells_li[1:]:
                pick, _ = pick_fa_for_fr_shift(day, fa_all_li, abs_fr, used, rng)
                ws[a1].value = pick or ""
                if pick:
                    used.add(pick)



# ---------------- OG assignment ----------------

FA_COUNTS: Dict[str, Dict[str,int]] = {d:{og:0 for og in OG_LIST} for d in WEEKDAYS}
def reset_og_counts():
    for d in WEEKDAYS:
        for og in OG_LIST: FA_COUNTS[d][og]=0

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
    Assign every present LA to all of their dedicated OGs (no prioritization),
    skipping 'Laufen' on days not in LAUFEN_DAYS.
    """
    reset_og_counts()

    for day in WEEKDAYS:
        abs_today = absences_by_day.get(day, set())

        for og in OG_LIST:
            if og == "Laufen" and day not in LAUFEN_DAYS:
                continue
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
    if s and s.role!=LA:
        s.og_nonleader_count += 1
        if s.role==AA: s.aa_og_count += 1
    if count_for_fa and s and s.role==OA:
        FA_COUNTS[day][og] += 1
    return True

def assign_nonleaders_to_ogs(ws: Worksheet, absences_by_day: Dict[str,Set[str]], seed: Optional[int]=None) -> Dict[str,Dict[str,int]]:
    rng = random.Random(seed)
    for day in WEEKDAYS:
        abs_today = absences_by_day.get(day,set())

        # 1) non-leader OAs with rotations (skip Laufen)
        for name in [n for n in (oa_bh+oa_li) if n not in abs_today]:
            s=staff_by_name[name]
            for og in s.rotations:
                if og in OG_LIST_NONLEADER: _place_in_og(ws,day,og,name,True)

        # 1b) AAs with rotations (skip Laufen)
        for name in [n for n in (aa_bh+aa_li) if n not in abs_today]:
            s=staff_by_name[name]
            for og in s.rotations:
                if og in OG_LIST_NONLEADER: _place_in_og(ws,day,og,name,False)

        # 2) no-rotation → lowest FA_COUNTS (skip Laufen)
        def free_ogs(day):
            """
            OGs that can receive 'free' assignments (no-rotation FAs/AAs).
            We explicitly *exclude* 'Mammo' so only:
              - LA leaders with leads_ogs={'Mammo'}
              - staff with rotation including 'Mammo'
            can ever be placed there.
            """
            return [
                og for og in OG_LIST_NONLEADER
                if og != "Mammo"
                and _first_empty_cell(ws, OG_CELLS[og][day]) is not None
                    ]
        for name in [n for n in (oa_bh+oa_li) if n not in abs_today and not staff_by_name[n].rotations]:
            opts = free_ogs(day); 
            if not opts: continue
            minv = min(FA_COUNTS[day][og] for og in opts); bucket=[og for og in opts if FA_COUNTS[day][og]==minv]
            og = rng.choice(bucket); _place_in_og(ws,day,og,name,True)

        for name in [n for n in (aa_bh+aa_li) if n not in abs_today and not staff_by_name[n].rotations]:
            opts = free_ogs(day); 
            if not opts: continue
            minv = min(FA_COUNTS[day][og] for og in opts); bucket=[og for og in opts if FA_COUNTS[day][og]==minv]
            og = rng.choice(bucket); _place_in_og(ws,day,og,name,False)

        # 3) coverage flags
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

def _fair_pick_pool(meeting_key: str, pool_idx: int, candidates: List[str], rng: random.Random) -> Optional[str]:
    """
    Fair pick for a given meeting pool:
      1) Prefer people with FEWER total meetings (meetings_count) over the week.
      2) Among those, prefer those with fewer uses in THIS meeting/pool (POOL_COUNTS).
    """
    if not candidates:
        return None

    # Step 1: global meeting load
    min_meet = min(staff_by_name[n].meetings_count for n in candidates)
    cand_by_meet = [n for n in candidates if staff_by_name[n].meetings_count == min_meet]

    # Step 2: per-meeting/pool fairness
    min_pool = min(POOL_COUNTS[(meeting_key, pool_idx, n)] for n in cand_by_meet)
    bucket = [n for n in cand_by_meet if POOL_COUNTS[(meeting_key, pool_idx, n)] == min_pool]

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

            pick = _fair_pick_pool(meeting_key, idx, cands, rng)
            if pick:
                _assign(ws, a1, pick, style)
                if monday_style and day == "Montag":
                    if monday_style == "red": set_red(ws, a1)
                _bump_pool(meeting_key, idx, pick)
                staff_by_name[pick].meetings_count += 1
                placed = True
                break

        if not placed and fallback_text is not None:
            _assign(ws, a1, fallback_text, fallback_style)

def assign_meetings(ws: Worksheet, absences: Dict[str, Set[str]], seed: Optional[int]=None):
    rng = random.Random(seed)

    write_medizin_placeholders_monday(ws)

    laufen_by_day = read_laufen_by_day(ws)
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

if __name__ == "__main__":
    import argparse, os

    parser = argparse.ArgumentParser(
        description="Run Wochenplan pipeline: OGs (LA→non-LA/AA) → FR → Meetings."
    )
    parser.add_argument("-i", "--input",  required=True,
                        help=".xlsm source file (e.g. KW 41_leer.xlsm)")
    parser.add_argument("-o", "--output", required=True,
                        help="FINAL .xlsm (after meetings), e.g. KW_41_FINAL.xlsm")
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

    # Read absences once (cleanup does not touch absence cells)
    absences = read_absences_by_day(ws)

    # 1) OG leaders (LA)
    assign_la_to_ogs(ws, absences)
    if args.keep_intermediate:
        wb.save(og_leaders_out)
        patch_xlsm(og_leaders_out, args.input)

    # 2) OG non-leaders (FA/AA) + flags
    assign_nonleaders_to_ogs(ws, absences, seed=args.seed)
    if args.keep_intermediate:
        wb.save(og_full_out)
        patch_xlsm(og_full_out, args.input)

    # 3) FR (depends on OG + Laufen in OG)
    assign_fr_shifts_to_cells(ws, absences, seed=args.seed)
    if args.keep_intermediate:
        wb.save(fr_out)
        patch_xlsm(fr_out, args.input)

    # 4) Meetings
    assign_meetings(ws, absences, seed=args.seed)

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
