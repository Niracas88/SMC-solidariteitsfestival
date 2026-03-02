"""
Solidariteitsfestival – Namiddagplanning (3 tijdssloten) – versie met nieuwe regels

REGELS
- 3 tijdssloten: Slot1, Slot2, Slot3
- Leerjaar 1/2/3: 1x Oeganda-getuigenis (vast per leerjaar), 1x workshop, 1x vrij moment
    * leerjaar 1 -> Oeganda in Slot1
    * leerjaar 2 -> Oeganda in Slot2
    * leerjaar 3 -> Oeganda in Slot3
- Leerjaar 4/5: gén Oeganda in namiddag, wel 1x info (Info 1..5), 1x workshop, 1x vrij moment
    * Info 1..5 kan in elk slot
    * Verdeel Info 1..5 zo gelijk mogelijk (globaal), geen cap
- Workshops (9) met capaciteit per workshop per slot:
    Mandala 45
    Yoga 14
    Bootcamp 150
    Dans 150
    Volksdans 30
    Lezen 15
    Mooimakers 24
    Schrijf ze vrij 26
    Hockey 40
- Restricties:
    * Schrijf ze vrij: enkel leerjaar 2 en 3
    * Hockey: enkel leerjaar 4 en 5
- Keuzes (top 3) gelden enkel voor workshops. Elke leerling krijgt precies 1 workshop:
    probeer keuze1 -> keuze2 -> keuze3; anders fallback naar workshop met plaats (en toegelaten voor dat leerjaar)
    PRIORITEIT: voorkeurkeuzes > klasgroepering
- NIEUWE REGEL: Zoveel mogelijk leerlingen van dezelfde klas samen in een activiteit (workshop/vrij/info)
    maar NIET ten koste van voorkeurkeuzes
- Output: Excel met tabs: planning, bezetting, kwaliteit, problemen, input_errors

VEREIST
    pip install pandas openpyxl numpy
"""

from __future__ import annotations
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple, Set

import numpy as np
import pandas as pd


# =========================
# CONFIG
# =========================
INPUT_XLSX = r"C:\Users\willem.vandenbussche\OneDrive - Stella Matutinacollege\WESP\python\Solidariteitsfestival_2026_opgekuist.xlsx"
# OUTPUT naar dezelfde map als het input bestand
OUTPUT_DIR = r"C:\Users\willem.vandenbussche\OneDrive - Stella Matutinacollege\WESP\python"
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "Solidariteitsfestival_2025_planning_output.xlsx")

SEED = 20250129

SLOTS = [1, 2, 3]
SLOT_COLS = {1: "Slot1", 2: "Slot2", 3: "Slot3"}

OEGANDA_LABEL = "Oeganda-getuigenis"
FREE_LABEL = "Vrij moment (Speelplaats Nieuwstraat)"
INFO_PREFIX = "Info "  # Info 1 .. Info 5

# Capaciteiten per slot
WORKSHOP_CAP = {
    "MANDALA": 45,
    "YOGA": 14,
    "BOOTCAMP": 150,
    "DANS": 150,
    "VOLKSDANS": 30,
    "LEZEN": 15,
    "MOOIMAKERS": 24,
    "SCHRIJF_ZE_VRIJ": 26,
    "HOCKEY": 40,
}
WORKSHOPS = list(WORKSHOP_CAP.keys())

# Excel-tekst -> code
WORKSHOP_MAP = {
    "Mandala": "MANDALA",
    "Mandala's": "MANDALA",
    "Yoga": "YOGA",
    "Mindfulness/yoga": "YOGA",
    "Bootcamp": "BOOTCAMP",
    "Dans": "DANS",
    "Volksdans": "VOLKSDANS",
    "Lezen": "LEZEN",
    "Mooimakers": "MOOIMAKERS",
    "Schrijf ze vrij": "SCHRIJF_ZE_VRIJ",
    "Hockey": "HOCKEY",
}

# Code -> weergavenaam (output)
WORKSHOP_LABEL = {
    "MANDALA": "Mandala",
    "YOGA": "Yoga",
    "BOOTCAMP": "Bootcamp",
    "DANS": "Dans",
    "VOLKSDANS": "Volksdans",
    "LEZEN": "Lezen",
    "MOOIMAKERS": "Mooimakers",
    "SCHRIJF_ZE_VRIJ": "Schrijf ze vrij",
    "HOCKEY": "Hockey",
}

# =========================
# HELPERS
# =========================

def make_rng(seed: int) -> np.random.Generator:
    return np.random.default_rng(seed)

def normalize_header(s: str) -> str:
    return " ".join(str(s).strip().lower().split())

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    norm_to_actual = {normalize_header(c): c for c in df.columns}

    def must_get(possible: List[str], label: str) -> str:
        for p in possible:
            k = normalize_header(p)
            if k in norm_to_actual:
                return norm_to_actual[k]
        raise ValueError(
            f"Ontbrekende kolom voor '{label}'.\n"
            f"Gezocht (een van): {possible}\n"
            f"Gevonden kolommen: {list(df.columns)}"
        )

    id_col = must_get(["id", "Id"], "Id")
    class_col = must_get(["klas", "Klas"], "Klas")
    c1 = must_get(["keuze 1", "eerste keuze", "keuze1"], "Keuze 1")
    c2 = must_get(["keuze 2", "tweede keuze", "keuze2"], "Keuze 2")
    c3 = must_get(["keuze 3", "derde keuze", "keuze3"], "Keuze 3")

    fn_col = None
    ln_col = None
    for poss, var in [
        (["voornaam", "first name", "firstname"], "voornaam"),
        (["naam", "familienaam", "last name", "lastname"], "naam"),
    ]:
        for p in poss:
            k = normalize_header(p)
            if k in norm_to_actual:
                if var == "voornaam":
                    fn_col = norm_to_actual[k]
                else:
                    ln_col = norm_to_actual[k]
                break

    return {"id": id_col, "klas": class_col, "c1": c1, "c2": c2, "c3": c3, "voornaam": fn_col, "naam": ln_col}

def grade_from_class(klas: str) -> Optional[int]:
    s = str(klas).strip()
    if not s:
        return None
    if s[0].isdigit():
        return int(s[0])
    return None

def normalize_choice(val: object) -> Optional[str]:
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return None
    up = s.upper()
    if up in WORKSHOP_CAP:
        return up
    # exact
    if s in WORKSHOP_MAP:
        return WORKSHOP_MAP[s]
    # case-insensitive
    for k, v in WORKSHOP_MAP.items():
        if k.strip().lower() == s.lower():
            return v
    return None

def preference_rank(prefs: List[str], w: str) -> int:
    try:
        return prefs.index(w) + 1
    except ValueError:
        return 999

def allowed_workshops_for_grade(grade: int) -> List[str]:
    allowed = WORKSHOPS.copy()
    if grade in (1, 4, 5):
        # Schrijf ze vrij niet voor 1/4/5
        if "SCHRIJF_ZE_VRIJ" in allowed:
            allowed.remove("SCHRIJF_ZE_VRIJ")
    if grade in (1, 2, 3):
        # Hockey niet voor 1/2/3
        if "HOCKEY" in allowed:
            allowed.remove("HOCKEY")
    # Voor 2/3 is Schrijf ze vrij wél toegestaan, voor 4/5 is Hockey wél toegestaan
    return allowed

def oeganda_slot_for_grade(grade: int) -> Optional[int]:
    if grade == 1:
        return 1
    if grade == 2:
        return 2
    if grade == 3:
        return 3
    return None  # 4/5 geen oeganda namiddag


# =========================
# DATA MODEL
# =========================

@dataclass
class Student:
    sid: int
    klas: str
    grade: int
    voornaam: str
    naam: str
    prefs: List[str]  # 3 workshop codes


# =========================
# CORE PLANNING
# =========================

def assign_infos_balanced(
    students_45: List[Student],
    rng: np.random.Generator,
) -> Tuple[Dict[int, int], Dict[int, str]]:
    """
    Wijs aan 4e/5e leerlingen:
    - 1 info-slot (1..3) zodat slots ongeveer gelijk gevuld zijn
    - 1 info-type (Info 1..5) zo gelijk mogelijk verdeeld (globaal)
    
    Prioriteert klasgroeperingen (secundair): leerlingen van dezelfde klas krijgen LIEVER hetzelfde info-slot
    maar dit mag voorkeurkeuzes NIET in gevaar brengen.
    
    Return:
      info_slot_by_id, info_label_by_id
    """
    # Groepeer per klas
    by_klas: Dict[str, List[Student]] = {}
    for s in students_45:
        if s.klas not in by_klas:
            by_klas[s.klas] = []
        by_klas[s.klas].append(s)

    info_slot_by_id: Dict[int, int] = {}
    info_label_by_id: Dict[int, str] = {}

    # Bereken target per slot
    n = len(students_45)
    base = n // 3
    rem = n % 3
    targets = [base + (1 if i < rem else 0) for i in range(3)]

    slot_counts = {1: 0, 2: 0, 3: 0}

    # Proces per klas om groeperingen te begunstigen (maar niet dwingen)
    klas_list = list(by_klas.keys())
    rng.shuffle(klas_list)

    for klas in klas_list:
        students_in_klas = by_klas[klas]
        # Kies slot met minste leerlingen, maar hou rekening met de
        # doelverdeling (`targets`); als een slot al op "target" zit mag
        # een andere slot gekozen worden.
        possible = [s for s in SLOTS if slot_counts[s] < targets[s - 1]]
        if possible:
            best_slot = min(possible, key=lambda s: slot_counts[s])
        else:
            # alle slots hebben de target bereikt – ga alsnog met de
            # minst gevulde verder (dit gebeurt slechts als klasgroottes
            # de balans licht verstoren)
            best_slot = min(SLOTS, key=lambda s: slot_counts[s])
        for s in students_in_klas:
            info_slot_by_id[s.sid] = best_slot
            slot_counts[best_slot] += 1

    # Verdeel info-types (1..5) globaal gelijk
    ids_sorted = sorted(students_45, key=lambda s: info_slot_by_id[s.sid])
    type_counts = {f"{INFO_PREFIX}{i}": 0 for i in range(1, 6)}
    
    for s in ids_sorted:
        types = list(type_counts.keys())
        rng.shuffle(types)
        types.sort(key=lambda x: type_counts[x])
        chosen = types[0]
        type_counts[chosen] += 1
        info_label_by_id[s.sid] = chosen

    return info_slot_by_id, info_label_by_id


def pick_workshop_for_student(
    s: Student,
    slot: int,
    cap_left: Dict[int, Dict[str, int]],
    allowed: List[str],
) -> Optional[str]:
    """
    Kies workshop voor student.
    PRIORITEIT:
    1. Voorkeurkeuzes (1, 2, 3) in volgorde
    2. Fallback: meeste vrije plaatsen
    
    Klasgroepering wordt NIET hier afgedwongen.
    """
    # This helper is still used in a few places and retained for backwards
    # compatibility, but the main assignment logic now lives in
    # ``assign_workshops`` below which handles first‑choice maximisation and
    # klasgroepering.  The behaviour here remains the original logic so that
    # it can be used for ad‑hoc lookups or unit tests if desired.

    # Stap 1: Voorkeuren (strikt)
    for w in s.prefs:
        if w in allowed and cap_left[slot][w] > 0:
            return w

    # Stap 2: Fallback - meeste vrije plaatsen
    available = [w for w in allowed if cap_left[slot][w] > 0]
    if not available:
        return None
    available.sort(key=lambda w: -cap_left[slot][w])
    return available[0]


def assign_free_slots(
    plan: Dict[int, Dict[int, Optional[str]]],
    students: List[Student],
    rng: np.random.Generator,
) -> None:
    """
    Wijs vrije momenten toe.
    Geen speciale klasgroepering hier (gewoon random uit lege slots).
    """
    by_id = {s.sid: s for s in students}

    for sid in by_id.keys():
        empty = [slot for slot in SLOTS if plan[sid][slot] is None]
        if len(empty) > 0:
            # Random keuze uit lege slots
            chosen_free = empty[int(rng.integers(0, len(empty)))]
            plan[sid][chosen_free] = FREE_LABEL


def build_plan(students: List[Student]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    rng = make_rng(SEED)
    ids = [s.sid for s in students]
    by_id = {s.sid: s for s in students}

    # capaciteit per slot
    cap_left = {slot: dict(WORKSHOP_CAP) for slot in SLOTS}

    # planning init (string labels)
    plan: Dict[int, Dict[int, Optional[str]]] = {sid: {slot: None for slot in SLOTS} for sid in ids}

    # 1) Oeganda vast invullen voor leerjaar 1/2/3
    for s in students:
        oslot = oeganda_slot_for_grade(s.grade)
        if oslot is not None:
            plan[s.sid][oslot] = OEGANDA_LABEL

    # 2) Info voor 4/5: kies info-slot + info-type (balanced)
    students_45 = [s for s in students if s.grade in (4, 5)]
    info_slot_by_id, info_label_by_id = assign_infos_balanced(students_45, rng)

    for s in students_45:
        islot = info_slot_by_id[s.sid]
        plan[s.sid][islot] = info_label_by_id[s.sid]

    # 3) Vrij moment: random uit lege slots
    assign_free_slots(plan, students, rng)

    # 4) Workshop: precies 1 per leerling, in het laatste lege slot
    #    De oorspronkelijke implementatie was een simpele lineaire toewijzing
    #    waarin leerlingen in willekeurige volgorde werden behandeld.  Daardoor
    #    kon een leerling later in de lus geen eerste keuze meer krijgen omdat
    #    een andere leerling die voorkeur eerder opeiste.  Bovendien werd
    #    klasgroepering onbenut gelaten.  De nieuwe routine hieronder verdeelt in
    #    fases: eerst alle eerste keuzes, dan tweede, derde en tenslotte een
    #    fallback.  Binnen elke fase lopen we klassen af, zodat leerlingen uit de
    #    zelfde klas bij voorkeur samen in hetzelfde lokaal terechtkomen.  De
    #    voorkeuren krijgen altijd voorrang; klasgroepering mag deze niet
    #    overrulen.

    def assign_workshops(
        students: List[Student],
        plan: Dict[int, Dict[int, Optional[str]]],
        cap_left: Dict[int, Dict[str, int]],
        rng: np.random.Generator,
    ) -> Dict[int, Optional[str]]:
        """
        Vul de lege plekken in ``plan`` met een workshopcode.

        Werkwijze per slot:
        1. bepaal alle leerlingen die nog niets hebben in dit slot
        2. doorloop voorkeurniveau 1..3
           * voor elk niveau loop je klas per klas (geshuffeld)
             - bij een leerling probeer je de workshop van dit niveau als
               capaciteit nog toelaatbaar is
        3. studenten die na de drie voorkeuren nog niets hebben, krijgen een
           fallback (meeste vrije plaatsen), waarbij we indien mogelijk een
           workshop kiezen waar al klasgenoten zitten

        Retourneert een dict ``workshop_assigned`` met de gekozen workshopcode of
        ``None`` als er geen plaats was.
        """

        workshop_assigned: Dict[int, Optional[str]] = {s.sid: None for s in students}
        ids = [s.sid for s in students]
        by_id_local = {s.sid: s for s in students}

        for slot in SLOTS:
            avail = [sid for sid in ids if plan[sid][slot] is None]

            # voorkeurstadia 1..3
            for stage in (0, 1, 2):
                # shuffle klassen zodat de volgorde niet deterministic is
                klas_order = list({by_id_local[sid].klas for sid in avail})
                rng.shuffle(klas_order)

                for klas in klas_order:
                    # verwerkt eerst alle leerlingen van deze klas
                    klas_sids = [sid for sid in avail if by_id_local[sid].klas == klas]
                    rng.shuffle(klas_sids)
                    for sid in klas_sids:
                        s = by_id_local[sid]
                        allowed = allowed_workshops_for_grade(s.grade)
                        pref = s.prefs[stage]
                        if pref in allowed and cap_left[slot][pref] > 0:
                            cap_left[slot][pref] -= 1
                            plan[sid][slot] = WORKSHOP_LABEL[pref]
                            workshop_assigned[sid] = pref
                            avail.remove(sid)

            # fallback voor overgebleven leerlingen
            if avail:
                klas_order = list({by_id_local[sid].klas for sid in avail})
                rng.shuffle(klas_order)
                for klas in klas_order:
                    klas_sids = [sid for sid in avail if by_id_local[sid].klas == klas]
                    rng.shuffle(klas_sids)
                    for sid in klas_sids:
                        s = by_id_local[sid]
                        allowed = allowed_workshops_for_grade(s.grade)
                        available = [w for w in allowed if cap_left[slot][w] > 0]
                        if not available:
                            plan[sid][slot] = "ONINGEPLAND_WORKSHOP"
                            workshop_assigned[sid] = None
                            avail.remove(sid)
                            continue

                        # probeer te kiezen voor een workshop waar klasgenoten
                        # al zitten in dit slot
                        classmates = [other for other in ids
                                      if other != sid
                                      and by_id_local[other].klas == s.klas
                                      and plan[other][slot] in WORKSHOP_LABEL.values()]
                        chosen = None
                        for w in available:
                            label = WORKSHOP_LABEL[w]
                            if any(plan[other][slot] == label for other in classmates):
                                chosen = w
                                break

                        if chosen is None:
                            available.sort(key=lambda w: -cap_left[slot][w])
                            chosen = available[0]

                        cap_left[slot][chosen] -= 1
                        plan[sid][slot] = WORKSHOP_LABEL[chosen]
                        workshop_assigned[sid] = chosen
                        avail.remove(sid)
        return workshop_assigned

    # voer de nieuwe toewijzing uit
    workshop_assigned = assign_workshops(students, plan, cap_left, rng)

    # 5) Output tabs
    planning_rows = []
    for s in students:
        planning_rows.append(
            {
                "Id": s.sid,
                "naam": s.naam,
                "voornaam": s.voornaam,
                "Klas": s.klas,
                "Leerjaar": s.grade,
                "Slot1": plan[s.sid][1],
                "Slot2": plan[s.sid][2],
                "Slot3": plan[s.sid][3],
            }
        )
    planning_df = pd.DataFrame(planning_rows)

    # bezetting
    bez_rows = []
    # workshops: we willen telling per SLOT per WORKSHOP_LABEL (outputnaam)
    for slot in SLOTS:
        # workshop usage per code
        for wcode in WORKSHOPS:
            used = WORKSHOP_CAP[wcode] - cap_left[slot][wcode]
            bez_rows.append(
                {
                    "Slot": slot,
                    "Activiteit": WORKSHOP_LABEL[wcode],
                    "Capaciteit": WORKSHOP_CAP[wcode],
                    "Toegewezen": used,
                    "Vrij": cap_left[slot][wcode],
                }
            )
        # labels
        bez_rows.append({"Slot": slot, "Activiteit": OEGANDA_LABEL, "Capaciteit": None,
                         "Toegewezen": sum(1 for sid in ids if plan[sid][slot] == OEGANDA_LABEL), "Vrij": None})
        bez_rows.append({"Slot": slot, "Activiteit": FREE_LABEL, "Capaciteit": None,
                         "Toegewezen": sum(1 for sid in ids if plan[sid][slot] == FREE_LABEL), "Vrij": None})
        for i in range(1, 6):
            lab = f"{INFO_PREFIX}{i}"
            bez_rows.append({"Slot": slot, "Activiteit": lab, "Capaciteit": None,
                             "Toegewezen": sum(1 for sid in ids if plan[sid][slot] == lab), "Vrij": None})
        bez_rows.append({"Slot": slot, "Activiteit": "ONINGEPLAND_WORKSHOP", "Capaciteit": None,
                         "Toegewezen": sum(1 for sid in ids if plan[sid][slot] == "ONINGEPLAND_WORKSHOP"), "Vrij": None})
    bezetting_df = pd.DataFrame(bez_rows)

    # kwaliteit: voorkeur-hit op de ene workshop per leerling + klasgroeperingsscore
    pref_counts = {1: 0, 2: 0, 3: 0, 999: 0}
    unplanned = 0
    
    # Klasgroepering: hoeveel % van leerlingen zit met klasgenoten in hetzelfde slot
    # Dit is een METING, niet een doel tijdens toewijzing
    class_cohesion_scores: Dict[str, float] = {}
    for klas in set(s.klas for s in students):
        students_in_klas = [s for s in students if s.klas == klas]
        if len(students_in_klas) == 0:
            continue
        
        total_cohesion = 0
        for slot in SLOTS:
            students_in_slot = [s for s in students_in_klas if plan[s.sid][slot] is not None]
            if len(students_in_slot) == 0:
                continue
            
            # Hoe veel zitten in dezelfde activiteit?
            activity_groups: Dict[Optional[str], int] = {}
            for s in students_in_slot:
                act = plan[s.sid][slot]
                activity_groups[act] = activity_groups.get(act, 0) + 1
            
            # Grootste groep in dit slot
            max_in_activity = max(activity_groups.values()) if activity_groups else 0
            total_cohesion += max_in_activity
        
        avg_cohesion = total_cohesion / len(students_in_klas) if students_in_klas else 0
        class_cohesion_scores[klas] = avg_cohesion

    avg_cohesion_overall = np.mean(list(class_cohesion_scores.values())) if class_cohesion_scores else 0
    
    for s in students:
        w = workshop_assigned[s.sid]
        if w is None:
            unplanned += 1
            continue
        r = preference_rank(s.prefs, w)
        if r not in (1, 2, 3):
            r = 999
        pref_counts[r] += 1

    kwaliteit_df = pd.DataFrame(
        [
            {"Metric": "Aantal leerlingen (geldig)", "Value": len(students)},
            {"Metric": "Workshop = keuze 1 (aantal)", "Value": pref_counts[1]},
            {"Metric": "Workshop = keuze 2 (aantal)", "Value": pref_counts[2]},
            {"Metric": "Workshop = keuze 3 (aantal)", "Value": pref_counts[3]},
            {"Metric": "Workshop = fallback (aantal)", "Value": pref_counts[999]},
            {"Metric": "ONINGEPLAND_WORKSHOP (aantal)", "Value": unplanned},
            {"Metric": "Gemiddelde klasgroepering (score 0-3)", "Value": f"{avg_cohesion_overall:.2f}"},
        ] + [
            {"Metric": f"Klasgroepering {klas}", "Value": f"{score:.2f}"}
            for klas, score in sorted(class_cohesion_scores.items())
        ]
    )

    # problemen: oningepland of slots niet correct gevuld
    prob_rows = []
    for s in students:
        vals = [plan[s.sid][1], plan[s.sid][2], plan[s.sid][3]]
        if "ONINGEPLAND_WORKSHOP" in vals or any(v is None for v in vals):
            # Veilig afhandelen van preferences (kunnen None bevatten)
            keuze1 = WORKSHOP_LABEL.get(s.prefs[0], "N/A") if s.prefs[0] is not None else "N/A"
            keuze2 = WORKSHOP_LABEL.get(s.prefs[1], "N/A") if s.prefs[1] is not None else "N/A"
            keuze3 = WORKSHOP_LABEL.get(s.prefs[2], "N/A") if s.prefs[2] is not None else "N/A"
            
            prob_rows.append(
                {
                    "Id": s.sid,
                    "naam": s.naam,
                    "voornaam": s.voornaam,
                    "Klas": s.klas,
                    "Leerjaar": s.grade,
                    "Slot1": vals[0],
                    "Slot2": vals[1],
                    "Slot3": vals[2],
                    "Keuze1": keuze1,
                    "Keuze2": keuze2,
                    "Keuze3": keuze3,
                }
            )
        # extra check: 1/2/3 moeten Oeganda hebben; 4/5 mogen niet
        oslot = oeganda_slot_for_grade(s.grade)
        if s.grade in (1, 2, 3) and (oslot is None or plan[s.sid][oslot] != OEGANDA_LABEL):
            prob_rows.append({"Id": s.sid, "Klas": s.klas, "Leerjaar": s.grade, "Probleem": "Oeganda ontbreekt of fout slot"})
        if s.grade in (4, 5) and OEGANDA_LABEL in vals:
            prob_rows.append({"Id": s.sid, "Klas": s.klas, "Leerjaar": s.grade, "Probleem": "Oeganda mag niet voor 4/5 in namiddag"})

    problemen_df = pd.DataFrame(prob_rows)

    return planning_df, bezetting_df, kwaliteit_df, problemen_df


# =========================
# MAIN
# =========================

def main() -> None:
    print("=" * 60)
    print("SOLIDARITEITSFESTIVAL PLANNING - START")
    print("=" * 60)
    
    # Check input file
    print(f"\n1. Checking input file: {INPUT_XLSX}")
    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(
            f"Bestand niet gevonden: {INPUT_XLSX}\n"
            f"Tip: zet INPUT_XLSX op een volledig pad."
        )
    print(f"   ✓ Input file found")
    
    # Read data
    print(f"\n2. Reading Excel data...")
    try:
        df = pd.read_excel(INPUT_XLSX)
        print(f"   ✓ Read {len(df)} rows")
    except Exception as e:
        print(f"   ✗ Error reading file: {e}")
        raise
    
    # Resolve columns
    print(f"\n3. Resolving column names...")
    try:
        cols = resolve_columns(df)
        print(f"   ✓ Columns resolved: {cols}")
    except Exception as e:
        print(f"   ✗ Error resolving columns: {e}")
        raise

    id_col = cols["id"]
    klas_col = cols["klas"]
    c1, c2, c3 = cols["c1"], cols["c2"], cols["c3"]
    fn_col = cols["voornaam"]
    ln_col = cols["naam"]

    # Normalize choices
    print(f"\n4. Normalizing workshop choices...")
    tmp = df.copy()
    tmp["__c1"] = tmp[c1].apply(normalize_choice)
    tmp["__c2"] = tmp[c2].apply(normalize_choice)
    tmp["__c3"] = tmp[c3].apply(normalize_choice)
    print(f"   ✓ Choices normalized")

    # Build student list
    print(f"\n5. Building student list...")
    students: List[Student] = []
    input_errors = []

    for idx, row in tmp.iterrows():
        try:
            sid = int(row[id_col])
            klas = str(row[klas_col]).strip()
            grade = grade_from_class(klas)
            if grade is None or grade not in (1, 2, 3, 4, 5):
                input_errors.append({"Id": sid, "Klas": klas, "Reden": "Leerjaar niet afleidbaar of buiten 1..5"})
                continue

            prefs = [row["__c1"], row["__c2"], row["__c3"]]
            if any(p is None for p in prefs):
                input_errors.append(
                    {
                        "Id": sid,
                        "Klas": klas,
                        "Leerjaar": grade,
                        "keuze 1 (raw)": row[c1],
                        "keuze 2 (raw)": row[c2],
                        "keuze 3 (raw)": row[c3],
                        "Reden": "Onbekende/lege workshopnaam in keuzes (update WORKSHOP_MAP indien nodig).",
                    }
                )
                continue

            voornaam = "" if fn_col is None else str(row[fn_col]).strip()
            naam = "" if ln_col is None else str(row[ln_col]).strip()

            students.append(Student(sid=sid, klas=klas, grade=grade, voornaam=voornaam, naam=naam, prefs=prefs))
        except Exception as e:
            print(f"   ✗ Error processing row {idx}: {e}")
            continue

    print(f"   ✓ {len(students)} valid students found")
    if len(input_errors) > 0:
        print(f"   ! {len(input_errors)} students with input errors")

    # Build plan
    print(f"\n6. Building schedule plan...")
    try:
        planning_df, bezetting_df, kwaliteit_df, problemen_df = build_plan(students)
        print(f"   ✓ Planning complete")
    except Exception as e:
        print(f"   ✗ Error building plan: {e}")
        raise

    input_errors_df = pd.DataFrame(input_errors)

    # Write output
    print(f"\n7. Writing Excel output: {OUTPUT_XLSX}")
    print(f"   Output directory: {OUTPUT_DIR}")
    
    # Check if output directory exists
    if not os.path.exists(OUTPUT_DIR):
        print(f"   ! Creating output directory: {OUTPUT_DIR}")
        os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    try:
        with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
            planning_df.to_excel(writer, index=False, sheet_name="planning")
            print(f"     - planning sheet: {len(planning_df)} rows")
            
            bezetting_df.to_excel(writer, index=False, sheet_name="bezetting")
            print(f"     - bezetting sheet: {len(bezetting_df)} rows")
            
            kwaliteit_df.to_excel(writer, index=False, sheet_name="kwaliteit")
            print(f"     - kwaliteit sheet: {len(kwaliteit_df)} rows")
            
            problemen_df.to_excel(writer, index=False, sheet_name="problemen")
            print(f"     - problemen sheet: {len(problemen_df)} rows")
            
            input_errors_df.to_excel(writer, index=False, sheet_name="input_errors")
            print(f"     - input_errors sheet: {len(input_errors_df)} rows")
        
        print(f"   ✓ Excel file written successfully")
    except Exception as e:
        print(f"   ✗ Error writing Excel file: {e}")
        raise

    # Final summary
    print(f"\n" + "=" * 60)
    print(f"✓ KLAAR")
    print(f"=" * 60)
    print(f"Output file: {OUTPUT_XLSX}")
    print(f"File exists: {os.path.exists(OUTPUT_XLSX)}")
    if os.path.exists(OUTPUT_XLSX):
        print(f"File size: {os.path.getsize(OUTPUT_XLSX)} bytes")
    
    if len(input_errors_df) > 0:
        print(f"\n⚠ Let op: {len(input_errors_df)} rijen met inputfouten (tab 'input_errors').")
    if len(problemen_df) > 0:
        print(f"⚠ Let op: {len(problemen_df)} rijen met planningproblemen (tab 'problemen').")
    print()


if __name__ == "__main__":
    main()
