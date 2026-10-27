"""
Budget Offer Automation Interface
Baker Hughes Thermodyn - DTI Department
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import tkinter.font as tkfont
from datetime import datetime
from PIL import Image, ImageTk
import os
import io, threading, time, json, math, re, shutil, subprocess, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────────────────────────────────────────────────────
# BAKER HUGHES UI COLOUR SYSTEM
# Green family and neutral greys only. Constant names are unchanged so the
# rest of the file needs no edit; the *_NAVY names now hold deep greens.
# ─────────────────────────────────────────────────────────────────────────────
# Brand greens
BH_GREEN   = "#018374"   # primary green: actions, active state, rules
BH_ACCENT  = "#02BC94"   # bright green: highlight on dark surfaces only
BH_GREEN_D = "#016358"   # pressed

# Deep greens (header, rail)
BH_DARK    = "#05322B"   # header, action bar
BH_NAVY    = "#0A3B33"   # rail, section headers
BH_NAVY_HI = "#12493F"   # hovered rail item
BH_NAVY_2  = "#083730"   # context strip
BH_LINE_D  = "#1D5247"   # rule on dark
BH_INK     = "#1E2423"   # primary text

# Secondary marks (no blue)
BH_TOOL    = "#3D4745"   # "built by the tool" status
BH_BLUE_LT = "#7FA39B"

# Grey and text
BH_GRAY    = "#4E5755"   # secondary text
BH_MUTED   = "#7B8482"   # muted text, axes
BH_GRID    = "#CBD0CE"
BH_BORDER  = "#DCE0DE"
BH_ON_DARK = "#8DB0A8"   # secondary text on dark
BH_ON_DARK2= "#CBDCD7"   # primary text on dark

# Surfaces
BH_BG      = "#F4F5F4"
BH_WHITE   = "#FFFFFF"
BH_LIGHT   = "#FAFBFA"
BH_SEC     = "#E5F1EE"   # selection tint
BH_TINT_B  = "#EDF0EF"

# Status
BH_OK      = "#1A6B3C"
BH_WARN    = "#B8860B"
BH_RED     = "#B22222"
BH_YELLOW  = "#B8860B"
BH_OK_BG   = "#E7F1EA"
BH_WARN_BG = "#F7EFDC"
BH_ERR_BG  = "#F6E7E7"
BH_TOOL_BG = "#EDF0EF"
BH_OK_PALE = "#A9C6B4"

# Progress fill: bright green to primary green
BH_G1      = "#02BC94"
BH_G2      = "#018374"
BH_G3      = "#018374"

APP_NAME  = "Y-B-TOOL"
APP_SUB   = "Budget Offer Generator"
APP_OWNER = "Baker Hughes · Thermodyn · DTI"
APP_VER   = "2.8"
SETTINGS_FILE = "y_b_tool_settings.json"

FONT_T  = ("Segoe UI", 22, "bold")
FONT_ST = ("Segoe UI", 12, "bold")
FONT_S  = ("Segoe UI", 10, "bold")
FONT_B  = ("Segoe UI", 9, "bold")
FONT_N  = ("Segoe UI", 9)
FONT_SM = ("Segoe UI", 8)
FONT_XS = ("Segoe UI", 7)
FONT_BD = ("Segoe UI", 8, "bold")
FONT_C  = ("Courier New", 9, "bold")
FONT_CS = ("Consolas", 8)

INCOTERMS = ["EXW","FCA","FAS","FOB","CFR","CIF","CPT","CIP","DAP","DPU","DDP"]
RATINGS   = ["150","300","600","900","1500","2500"]
APPS      = ["Air Service","Ammonia Synthesis","Ethylene Synthesis","FCC","Fuel Gas",
             "LNG","LPG","Lubricant Production","Methanol Synthesis","Olefin","Other",
             "Oxygen Service","Pipeline","Reforming","Reinjection","Urea Synthesis"]

PLC_DOCS = {
    "Siemens Common PLC": {
        "Base Case":          ["RPLIC004TOSIEM","RNTIC053ICLSIE","RPLIC006ICLSIE", "RPLIC009ICLSIE"],
        "Intrinsically Safe": ["RPLIC004TOSIEM","RNTIC053ICLSIE","RPLIC006ICLSIEIS","RPLIC009ICLSIEIS"],
        "Surge Arrestor":     ["RPLIC004TOSIEM","RNTIC053ICLSIE","RPLIC006ICLSIESA","RPLIC009ICLSIESA"],
    },
    "Siemens Separate PLC": {
        "Base Case":          ["RPLIC104TOSIEM","RNTIC053ICLSIE","RPLIC106ICLSIE", "RPLIC109ICLSIE"],
        "Intrinsically Safe": ["RPLIC104TOSIEM","RNTIC053ICLSIE","RPLIC106ICLSIEIS","RPLIC109ICLSIEIS"],
        "Surge Arrestor":     ["RPLIC104TOSIEM","RNTIC053ICLSIE","RPLIC106ICLSIESA","RPLIC109ICLSIESA"],
    },
    "Allen Bradley Common PLC": {
        "Base Case":          ["RPLIC004TOROCK","RNTIC153ICLROC","RPLIC006ICLROC", "RPLIC009ICLROC"],
        "Intrinsically Safe": ["RPLIC004TOROCK","RNTIC153ICLROC","RPLIC006ICLROCIS","RPLIC009ICLROCIS"],
        "Surge Arrestor":     ["RPLIC004TOROCK","RNTIC153ICLROC","RPLIC006ICLROCSA","RPLIC009ICLROCSA"],
    },
    "Allen Bradley Separate PLC": {
        "Base Case":          ["RPLIC104TOROCK","RNTIC153ICLROC","RPLIC106ICLROC", "RPLIC109ICLROC"],
        "Intrinsically Safe": ["RPLIC104TOROCK","RNTIC153ICLROC","RPLIC106ICLROCIS","RPLIC109ICLROCIS"],
        "Surge Arrestor":     ["RPLIC104TOROCK","RNTIC153ICLROC","RPLIC106ICLROCSA","RPLIC109ICLROCSA"],
    },
}

CASINGS = ["300", "350", "400", "450", "500", "600", "700", "800"]


def _casing_digits(value):
    """Casing as it appears in the machine name: 450 -> 45, 45 -> 45."""
    d = re.sub(r"\D", "", str(value or ""))
    if len(d) == 3 and d.endswith("0"):
        return d[:2]
    return d[:2]


COOLING_CONCEPTS = [
    "multistage_back_to_back_with_fan",
    "multistage_back_to_back_with_interstage_cooler_and_TCV",
    "multistage_back_to_back_with_TCV",
    "multistage_with_fan",
    "multistage_with_TCV",
    "single_stage",
    "single_stage_with_cooler",
]


def _flat(text):
    """Lower case, no accent, letters and digits only: how names are compared."""
    import unicodedata
    t = unicodedata.normalize("NFKD", str(text or ""))
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", t.lower())


def _concept_features(text):
    n = _flat(text)
    return dict(b2b=("backtoback" in n or "b2b" in n),
                single="single" in n,
                fan="fan" in n,
                tcv=("tcv" in n or "piquage" in n),
                interstage="interstage" in n,
                cooler="cooler" in n)


def _concept_for(family, cooling):
    """family: 'b2b' | 'multi' | 'single'; cooling: 'Fan' | 'TCV'."""
    if family == "single":
        return "single_stage"
    head = "multistage_back_to_back" if family == "b2b" else "multistage"
    return f"{head}_with_{'fan' if cooling == 'Fan' else 'TCV'}"


def _concept_family(name):
    f = _concept_features(name)
    return "single" if f["single"] else "b2b" if f["b2b"] else "multi"


# Incoterms offered in a proposal. EXW and DDP are never offered.
INCOTERMS = ["FCA", "FOB", "CFR", "CPT", "CIF", "CIP", "FAS", "DAP", "DPU"]
INCOTERMS_RECOMMENDED = ("FCA", "FOB", "CFR", "CPT", "CIF", "CIP")
INSURANCE_NOTE_RULES = ("FAS", "DAP", "DPU")
INSURANCE_NOTE = ("{code} does not oblige the buyer nor the seller to insure the "
                  "goods: insurance is therefore to be addressed elsewhere in the "
                  "sales contract.")


def incoterm_note(code):
    return INSURANCE_NOTE.format(code=code) if code in INSURANCE_NOTE_RULES else ""


# ─────────────────────────────────────────────────────────────────────────────
# STANDARD DOCUMENT LIBRARY (DB-2026)
#
# One folder per document, named after it. Inside, the tool takes the file
# that fits the offer: the concept of the cooling system, the P&ID of that
# concept, the GAD nearest to the machine, the PLC documents of the brand
# and option, every comment-and-exception file…
#
#   rule  single   the best file of the folder (newest, PDF first)
#         all      every file, sub-folder by sub-folder
#         concept  the file named after the cooling concept
#         pid      the P&ID drawn for that concept
#         gad      the GAD named XICLXXX_MGVX nearest to the machine
#         plc      the documents listed for the PLC brand and option
#         project  supplied by the engineer in the project folder
#         tool     built by this tool
# ─────────────────────────────────────────────────────────────────────────────
LIB_EXT = (".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".pptx", ".ppt",
           ".png", ".jpg", ".jpeg", ".tif", ".tiff")
LIB_NAME_RX = re.compile(r"^db[\s_-]*(20\d\d)$", re.I)


def _L(id_, name, title, aliases, rule="single", prefer=(), on=(False, False, True)):
    return dict(id=id_, name=name, title=title,
                aliases=[_flat(a) for a in aliases if _flat(a)],
                rule=rule, prefer=tuple(_flat(p) for p in prefer),
                on=dict(zip(("DTI", "BUDGET", "FIRM"), on)))


ALWAYS = (True, True, True)
DB_LIBRARY = [
    _L("cooling", "Cooling_Concept", "Cooling System Concept",
       ("cooling concept", "cooling concepts", "cooling"), rule="concept", on=ALWAYS),
    _L("gad", "GAD", "General Arrangement Drawing (GAD)",
       ("gad", "general arrangement"), rule="gad", on=ALWAYS),
    _L("pid", "P&ID", "P&ID", ("p&id", "pids", "p and id"), rule="pid", on=ALWAYS),
    _L("scope", "Scope_of_Supply", "Scope of Supply", ("scope of supply",),
       rule="tool", on=ALWAYS),
    _L("plc", "Associated_PLC_Documents", "PLC Documents",
       ("associated plc documents", "plc documents", "plc"), rule="plc"),
    _L("amb", "AMB_panels", "AMB Panels", ("amb panels", "amb panel", "amb"),
       prefer=("RNTICITO300V2",)),
    _L("cables", "cables", "Cables", ("cables", "cable")),
    _L("gepc", "GEPC_TECH_Offer", "GEPC Technical Offer",
       ("gepc tech offer", "gepc"), rule="project"),
    _L("utility", "Utility_and_Load_List", "Utility and Load List",
       ("utility and load list", "utility load list", "utility")),
    _L("safety", "ICL_Safety_Concept", "ICL Safety Concept",
       ("icl safety concept", "safety concept")),
    _L("landing", "Landing_Bearing_Philosophy", "Landing Bearing Philosophy",
       ("landing bearing philosophy", "landing bearing", "landing")),
    _L("hookup", "Hook-Up", "Hook-Up Drawings", ("hook-up", "hook up")),
    _L("piping", "Piping_Class", "Piping Class", ("piping class", "piping")),
    _L("painting", "Painting_Spec", "Painting Specification",
       ("painting spec", "painting specification", "painting")),
    _L("packing", "Packing_Specification", "Packing Specification",
       ("packing specification", "packing")),
    _L("qcp", "QCP", "Quality Control Plan (QCP)", ("qcp",)),
    _L("dilution", "Dilution_Sequence", "Dilution Sequence",
       ("sequence de dilution", "dilution sequence", "dilution")),
    _L("spare", "Spare_Part_List", "Spare Part List",
       ("spare part list", "spare parts list", "spare parts", "spare")),
    _L("tools", "Special_Tools_Doc", "Special Tools",
       ("special tools doc", "special tools")),
    _L("subsup", "Sub_Supplier_List", "Sub-Supplier List",
       ("sub supplier list", "sub suppliers", "sub supplier", "vendor list")),
    _L("trs", "TRS__CSE_Documentation", "Technical and Regulation Standards",
       ("trs cse documentation", "trs cse", "trs")),
    _L("ce", "C&E's", "Comments and Exceptions",
       ("c&e's", "c&es", "c&e's examples", "c&e examples"), rule="all"),
    _L("vdr", "VDR", "Vendor Document Register (VDR)", ("vdr",)),
    _L("reflist", "ICL_Reference_List", "ICL Reference List",
       ("icl reference list", "icl ref list", "reference list", "ref list")),
]
DB_BY_ID = {d["id"]: d for d in DB_LIBRARY}


def find_library(*near):
    """The newest DB-20xx folder next to the tool, above it, or above the
    project folder."""
    seen, hits = set(), []
    for base in near:
        if not base:
            continue
        d = os.path.abspath(base)
        for _ in range(4):
            for cand in (d, *[os.path.join(d, n) for n in _safe_listdir(d)]):
                name = os.path.basename(cand)
                m = LIB_NAME_RX.match(name)
                if m and os.path.isdir(cand) and cand not in seen:
                    seen.add(cand); hits.append((int(m.group(1)), cand))
            parent = os.path.dirname(d)
            if parent == d:
                break
            d = parent
    return max(hits)[1] if hits else ""


def _safe_listdir(d):
    try:
        return sorted(os.listdir(d))
    except OSError:
        return []


def lib_folders(root):
    """{document id: folder}. Folder names are compared without case, accent
    or punctuation; a sub-folder one level down is accepted too."""
    dirs = []
    for n in _safe_listdir(root):
        p = os.path.join(root, n)
        if os.path.isdir(p):
            dirs.append((n, p))
            for m in _safe_listdir(p):
                q = os.path.join(p, m)
                if os.path.isdir(q):
                    dirs.append((m, q))
    out, taken = {}, set()
    for exact in (True, False):
        for doc in DB_LIBRARY:
            if doc["id"] in out:
                continue
            best, score = None, 0
            for n, p in dirs:
                if p in taken:
                    continue
                f = _flat(n)
                names = [_flat(doc["name"])] + doc["aliases"]
                if exact:
                    s = 3 if f in names else 0
                else:
                    s = max([2 if f.startswith(a) and len(a) >= 3 else
                             1 if len(a) >= 5 and a in f else 0 for a in names])
                if s > score:
                    best, score = p, s
            if best:
                out[doc["id"]] = best
                taken.add(best)
    return out


def _lib_files(folder):
    out = []
    if not folder or not os.path.isdir(folder):
        return out
    for dp, dns, fns in os.walk(folder):
        dns[:] = sorted(d for d in dns if not d.startswith("."))
        for fn in sorted(fns):
            if fn.startswith(("~$", ".")) or not fn.lower().endswith(LIB_EXT):
                continue
            out.append(os.path.join(dp, fn))
    return out


def _rank_files(files, prefer=()):
    """Best file first: preferred name, then PDF, then the newest."""
    def key(p):
        n = _flat(os.path.basename(p))
        ext = os.path.splitext(p)[1].lower()
        s = 10 * any(x in n for x in prefer)
        s += {".pdf": 3, ".docx": 2, ".xlsx": 2, ".xlsm": 2}.get(ext, 1)
        s -= 5 * any(x in n for x in ("old", "draft", "copy", "backup", "obsolete"))
        try:
            age = os.path.getmtime(p)
        except OSError:
            age = 0
        return (-s, -age)
    return sorted(files, key=key)


def _one_per_stem(files):
    """The same document saved as PDF and Word: keep the PDF."""
    best = {}
    for p in files:
        k = (os.path.dirname(p), _flat(os.path.splitext(os.path.basename(p))[0]))
        if k not in best or _rank_files([p, best[k]])[0] == p:
            best[k] = p
    return [p for p in files if best.get(
        (os.path.dirname(p), _flat(os.path.splitext(os.path.basename(p))[0]))) == p]


FEATURE_WEIGHTS = (("b2b", 3, "back-to-back"), ("single", 3, "single stage"),
                   ("fan", 2, "fan"), ("tcv", 2, "TCV"),
                   ("interstage", 1, "interstage cooler"), ("cooler", 1, "cooler"))


def _feature_fit(target, name):
    t, f = _concept_features(target), _concept_features(name)
    score, diffs = 0, []
    for k, w, label in FEATURE_WEIGHTS:
        if t[k] == f[k]:
            score += w
        else:
            score -= w
            diffs.append(("with " if f[k] else "no ") + label)
    return score, diffs


GAD_MODEL = re.compile(r"(?<![a-z])(\d)?\s*icl[\s_-]?(\d{2})(\d{1,2})(?!\d)", re.I)
GAD_MOTOR = re.compile(r"mgv[\s_-]?(\d)", re.I)


def _gad_fit(name, ctx):
    m = GAD_MODEL.search(name)
    if not m:
        return None
    phases = m.group(1) or "1"
    casing, imp = m.group(2), int(m.group(3))
    mm = GAD_MOTOR.search(name)
    motor = mm.group(1) if mm else ""
    dist, diffs = 0, []
    want_ph = "2" if ctx.get("phases") == "2" else "1"
    if phases != want_ph:
        dist += 100
        diffs.append("back-to-back" if phases == "2" else "single casing")
    want_cs = ctx.get("casing") or ""
    if want_cs and casing != want_cs:
        try:
            dist += 4 * abs(int(casing) - int(want_cs))
        except ValueError:
            dist += 20
        diffs.append(f"casing {casing}0")
    try:
        want_imp = int(ctx.get("impellers") or 0)
    except ValueError:
        want_imp = 0
    if want_imp and imp != want_imp:
        dist += 6 * abs(imp - want_imp)
        diffs.append(f"{imp} impellers")
    want_mot = re.sub(r"\D", "", ctx.get("motor") or "")
    if want_mot and motor and motor != want_mot:
        dist += 40
        diffs.append(f"MGV{motor}")
    return dist, diffs


def lib_resolve(doc, folder, ctx, choice=None):
    """What the library gives for one document.

    Returns dict(files, quality, note, candidates): quality is exact,
    closest, chosen, project, tool or missing; candidates are the other
    files that could stand in, best first, each with a short reason."""
    out = dict(files=[], quality="missing", note="", candidates=[], folder=folder)
    rule = doc["rule"]
    if rule == "tool":
        out.update(quality="tool", note="built by the tool")
        return out
    if rule == "project":
        mine = [p for p in ctx.get("project_files", [])
                if any(a in _flat(os.path.basename(p)) for a in doc["aliases"]
                       if len(a) >= 4)]
        ranked = _rank_files(mine) + _rank_files(_lib_files(folder))
        out["candidates"] = [(p, "") for p in ranked]
        if choice and os.path.exists(choice):
            out.update(files=[choice], quality="chosen", note="chosen by hand")
        elif mine:
            out.update(files=[ranked[0]], quality="project",
                       note="found in the project folder")
        elif ranked:
            out.update(files=[ranked[0]], quality="exact", note="from the library")
        else:
            out["note"] = "to be supplied by the engineer"
        return out
    if not folder:
        out["note"] = f"no {doc['name']} folder in the library"
        return out
    files = _lib_files(folder)
    if not files:
        out["note"] = "the folder is empty"
        return out

    if rule == "all":
        picked = _one_per_stem(files)
        out.update(files=picked, quality="exact",
                   note=f"{len(picked)} document(s)",
                   candidates=[(p, "") for p in picked])
        return out

    if rule == "plc":
        ucp, opt = ctx.get("ucp", ""), ctx.get("plc_option", "")
        brand = ("siemens",) if "siemens" in _flat(ucp) else \
            ("rockwell", "allenbradley", "allen")
        sub = [d for d in sorted({os.path.dirname(p) for p in files})
               if any(b in _flat(os.path.relpath(d, folder)) for b in brand)]
        pool = [p for p in files if any(p.startswith(d + os.sep) for d in sub)] \
            if sub else files
        codes = PLC_DOCS.get(ucp, {}).get(opt, [])
        picked, missing = [], []
        for code in codes:
            cf = _flat(code)
            best, rank = None, None
            for p in _rank_files(pool):
                s = _flat(os.path.splitext(os.path.basename(p))[0])
                i = s.find(cf)
                if i < 0:
                    continue
                tail = s[i + len(cf):]
                r = 0 if tail == "" else 2 if tail[:2] in ("is", "sa") else 1
                if rank is None or r < rank:
                    best, rank = p, r
            if best and best not in picked:
                picked.append(best)
            elif not best:
                missing.append(code)
        out["candidates"] = [(p, "") for p in _one_per_stem(pool)]
        if picked:
            out.update(files=picked, quality="exact" if not missing else "closest",
                       note=(f"{len(picked)} document(s)"
                             + (f", missing {', '.join(missing)}" if missing else "")))
        else:
            fall = _one_per_stem(pool)
            out.update(files=fall, quality="closest",
                       note="document numbers not found in the file names: every "
                            "file of the brand folder is taken")
        return out

    if rule == "concept":
        target = (ctx.get("concept") or "").strip()
        out["candidates"] = [(p, "") for p in _rank_files(files)]
        exact = [p for p in files
                 if os.path.splitext(os.path.basename(p))[0].strip().lower()
                 == target.lower()]
        if exact:
            out.update(files=[_rank_files(exact)[0]], quality="exact", note="")
        else:
            out["note"] = f"no file named {target} in the folder"
        if choice and os.path.exists(choice):
            out.update(files=[choice], quality="chosen", note="chosen by hand")
        return out

    if rule == "pid":
        target = ctx.get("concept", "")
        scored = []
        for p in files:
            stem = os.path.splitext(os.path.basename(p))[0]
            if rule == "concept" and _flat(stem) == _flat(target):
                scored.append((100, [], p))
                continue
            s, diffs = _feature_fit(target, stem)
            scored.append((s, diffs, p))
        top = max(s for s, _d, _p in scored)
        best = _rank_files([p for s, _d, p in scored if s == top])
        ranked = sorted(scored, key=lambda x: (-x[0], _rank_files(files).index(x[2])))
        out["candidates"] = [(p, ", ".join(d)) for _s, d, p in ranked]
        pick = best[0]
        diffs = next(d for s, d, p in scored if p == pick)
        exact = top == 100 or not diffs
        out.update(files=[pick], quality="exact" if exact else "closest",
                   note="" if exact else
                   f"no exact match for {target}: nearest file is "
                   + ", ".join(diffs))
    elif rule == "gad":
        fits = [(f, p) for p in files
                for f in [_gad_fit(os.path.basename(p), ctx)] if f is not None]
        if not fits:
            ranked = _rank_files(files)
            out.update(files=[ranked[0]], quality="closest",
                       note="no file named like XICLXXX_MGVX",
                       candidates=[(p, "") for p in ranked])
        else:
            fits.sort(key=lambda x: (x[0][0], _rank_files(files).index(x[1])))
            out["candidates"] = [(p, ", ".join(d) or "exact") for (_dist, d), p in fits]
            (dist, diffs), pick = fits[0]
            want = (f"{'2' if ctx.get('phases') == '2' else ''}ICL"
                    f"{ctx.get('casing', '')}{ctx.get('impellers', '')}"
                    f"_{ctx.get('motor', '')}")
            out.update(files=[pick], quality="exact" if dist == 0 else "closest",
                       note="" if dist == 0 else
                       f"no GAD for {want}: nearest has " + ", ".join(diffs))
    else:                                               # single
        ranked = _rank_files(files, doc["prefer"])
        out.update(files=[ranked[0]], quality="exact",
                   note=f"{len(files)} files in the folder: newest taken"
                   if len(files) > 1 else "",
                   candidates=[(p, "") for p in ranked])

    if choice and os.path.exists(choice) and rule != "all":
        out.update(files=[choice], quality="chosen", note="chosen by hand")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# CONTACTS — read from the Excel list kept next to the tool
#
#   row n     Sales manager | Sales advisor | Application engineer   (merged)
#   row n+1   Name  Phone  Email  [Address] [Title] under each of them
#   below     one person per row
# ─────────────────────────────────────────────────────────────────────────────
CONTACT_ROLES = (("sales", r"sales\s*manager"),
                 ("adv", r"sales\s*advis"),
                 ("ae", r"application\s*eng|^\s*ae\s*$"))
CONTACT_FIELDS = (("name", r"name|nom"), ("phone", r"phone|t[ée]l|mobile|gsm"),
                  ("email", r"mail"), ("address", r"addr|adres"),
                  ("title", r"title|position|job|fonction|function"))


def read_contacts(path):
    """{role: [dict(name, phone, email, address, title)]} from the workbook."""
    found = {r: [] for r, _rx in CONTACT_ROLES}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    for ws in wb.worksheets:
        grid = [[("" if v is None else str(v).strip()) for v in row]
                for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400),
                                        values_only=True)]
        for r in range(min(len(grid) - 1, 30)):
            heads = []
            for c, v in enumerate(grid[r]):
                for role, rx in CONTACT_ROLES:
                    if v and re.search(rx, v, re.I):
                        heads.append((c, role))
            if not heads:
                continue
            width = max(len(x) for x in grid)
            spans = []
            for i, (c, role) in enumerate(heads):
                end = heads[i + 1][0] if i + 1 < len(heads) else width
                spans.append((role, c, end))
            fields_row = grid[r + 1]
            for role, c0, c1 in spans:
                cols = {}
                for c in range(c0, min(c1, len(fields_row))):
                    for key, rx in CONTACT_FIELDS:
                        if key not in cols and fields_row[c] and \
                                re.search(rx, fields_row[c], re.I):
                            cols[key] = c
                            break
                if "name" not in cols:
                    continue
                blank = 0
                for row in grid[r + 2:]:
                    rec = {k: (row[c] if c < len(row) else "") for k, c in cols.items()}
                    if not rec.get("name"):
                        blank += 1
                        if blank >= 3:
                            break
                        continue
                    blank = 0
                    rec["phone"] = rec.get("phone", "").replace(".0", "") \
                        if rec.get("phone", "").endswith(".0") else rec.get("phone", "")
                    found[role].append(rec)
            if any(found.values()):
                return found
    return found


def find_contacts_file(folder, hint=""):
    if hint and os.path.exists(hint):
        return hint
    names = [n for n in _safe_listdir(folder)
             if n.lower().endswith((".xlsx", ".xlsm")) and not n.startswith("~$")]
    names.sort(key=lambda n: (0 if "contact" in n.lower() else 1, n.lower()))
    for n in names:
        p = os.path.join(folder, n)
        try:
            if any(read_contacts(p).values()):
                return p
        except Exception:
            continue
    return ""


def write_contacts_template(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contacts"
    thin = Side(style="thin", color="7B8482")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, role in enumerate(("Sales manager", "Sales advisor",
                              "Application engineer")):
        c0 = 1 + 3 * i
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 2)
        ws.cell(1, c0, role).alignment = Alignment(horizontal="center")
        for j, f in enumerate(("Name", "Phone", "Email")):
            cell = ws.cell(2, c0 + j, f)
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[openpyxl.utils.get_column_letter(c0 + j)].width = 24
        for r in range(1, 8):
            for j in range(3):
                ws.cell(r, c0 + j).border = box
    wb.save(path)


def login_tokens():
    names = [os.environ.get(k, "") for k in ("USERNAME", "USER", "LOGNAME")]
    try:
        names.append(os.getlogin())
    except Exception:
        pass
    return [_flat(n) for n in names if n]


def contact_matches_login(rec, tokens):
    parts = re.split(r"[\s.@_-]+", f"{rec.get('name', '')} "
                     f"{(rec.get('email') or '').split('@')[0]}")
    parts = [_flat(p) for p in parts if len(_flat(p)) >= 4]
    return sum(1 for t in tokens for p in parts if p in t or t in p)


# ─────────────────────────────────────────────────────────────────────────────
# COMMERCIAL SECTION — the Word file kept in Section 11_Commercial of the
# project folder, with its price, delivery time and Incoterm brought up to date
# ─────────────────────────────────────────────────────────────────────────────
COMMERCIAL_LABELS = (
    ("price", r"(?:base|total|selling)\s+price"),
    ("delivery", r"delivery\s+time|lead\s+time"),
    ("incoterm", r"delivery\s+place(?:\s*\(\s*incoterms?[^)]*\))?|incoterms?"),
)
INCOTERM_TOKEN = re.compile(
    r"\b(ex[\s-]?works|exw|fca|fas|fob|cfr|cif|cpt|cip|dap|dpu|ddp)\b", re.I)
AMOUNT_RX = re.compile(r"\d{1,3}(?:[,.\u00a0\u202f ]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?")


def fill_commercial_docx(src, dst, values):
    """Copy src to dst with the values of the offer written after their
    labels. values: price, delivery (a number), incoterm, insurance_note.
    The value may sit in the same paragraph as its label or in the next
    cell of a table. Returns report lines (field, what happened)."""
    with zipfile.ZipFile(src) as z:
        names = [i.filename for i in z.infolist()]
        files = {n: z.read(n) for n in names}
    doc = files["word/document.xml"].decode("utf-8")
    par = re.compile(r"<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>", re.S)
    report, pending, done = [], {}, set()
    pieces, cur, note_done = [], 0, False

    colon_seen = [True]

    def put(x, field, start):
        """Replace the value of field found from start; (new xml, old text)."""
        t = _dx_text_raw(x)
        val = (values.get(field) or "").strip()
        if not val:
            return x, None
        if field == "price":
            m = AMOUNT_RX.search(t, start)
            if m:
                # an amount also written in words follows the figure
                w = re.match(r"\s*\(([a-z ,-]+)\)", t[m.end():], re.I)
                if w and re.search(r"million|thousand|hundred", w.group(1), re.I):
                    words = _num_words(int(round(_num(val))))
                    if "euro" in w.group(1).lower():
                        words += " euros"
                    a = m.end() + w.start(1)
                    x = _dx_sub(x, a, a + len(w.group(1)), words)
                    t = _dx_text_raw(x)
                    m = AMOUNT_RX.search(t, start)
        elif field == "delivery":
            m = re.compile(r"\d+").search(t, start)
        else:
            m = INCOTERM_TOKEN.search(t, start)
            if not m:
                rest = t[start:].strip()
                if not rest or not colon_seen[0]:
                    return x, None
                a = t.index(rest, start)
                return _dx_sub(x, a, a + len(rest), val), rest
        if not m:
            return x, None
        return _dx_sub(x, m.start(), m.end(), val), m.group(0)

    for m in par.finditer(doc):
        x = m.group(0)
        t = _dx_text_raw(x)
        new, hit_incoterm = x, False
        labelled = False
        for field, rx in COMMERCIAL_LABELS:
            lm = re.search(rx, t, re.I)
            if not lm:
                continue
            labelled = True
            colon = t.find(":", lm.end())
            colon_seen[0] = 0 <= colon - lm.end() < 25
            start = colon + 1 if colon_seen[0] else lm.end()
            new, old = put(new, field, start)
            colon_seen[0] = True
            if old is not None:
                report.append((field, f"{old.strip()}  →  {values[field]}"))
                done.add(field)
                hit_incoterm = field == "incoterm"
            elif (values.get(field) or "").strip():
                pending[field] = 3              # value in the next cell
            break
        if not labelled and t.strip():
            for field in list(pending):
                new, old = put(new, field, 0)
                if old is not None:
                    report.append((field, f"{old.strip()}  →  {values[field]}"))
                    done.add(field)
                    hit_incoterm = field == "incoterm"
                    del pending[field]
                    break
                pending[field] -= 1
                if pending[field] <= 0:
                    del pending[field]
        pieces.append(doc[cur:m.start()]); pieces.append(new); cur = m.end()
        note = values.get("insurance_note") or ""
        if note and hit_incoterm and not note_done:
            sp = _dx_ppr(new)
            ppr = new[sp[0]:sp[1]] if sp else ""
            ppr = re.sub(r"<w:numPr>.*?</w:numPr>", "", ppr, flags=re.S)
            pieces.append("<w:p>" + ppr + "<w:r><w:rPr><w:i/></w:rPr>"
                          '<w:t xml:space="preserve">'
                          + _html.escape(note, quote=False) + "</w:t></w:r></w:p>")
            note_done = True
    pieces.append(doc[cur:])
    files["word/document.xml"] = "".join(pieces).encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, files[n])
    labels = {"price": "base price", "delivery": "delivery time",
              "incoterm": "delivery place (Incoterm)"}
    for field, _rx in COMMERCIAL_LABELS:
        if field not in done and (values.get(field) or "").strip():
            report.append((field, f"{labels[field]}: label not found in the document"))
    if note_done:
        report.append(("incoterm", "insurance note added under the Incoterm"))
    return report


def _num_words(n):
    """4600000 -> four million six hundred thousand."""
    ones = ("zero one two three four five six seven eight nine ten eleven twelve "
            "thirteen fourteen fifteen sixteen seventeen eighteen nineteen").split()
    tens = ("", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
            "eighty", "ninety")

    def small(k):
        out = []
        if k >= 100:
            out += [ones[k // 100], "hundred"]; k %= 100
        if k >= 20:
            out.append(tens[k // 10] + ("-" + ones[k % 10] if k % 10 else ""))
        elif k:
            out.append(ones[k])
        return " ".join(out)
    if n <= 0:
        return "zero"
    parts = []
    for size, name in ((10**9, "billion"), (10**6, "million"),
                       (10**3, "thousand"), (1, "")):
        if n >= size:
            parts.append((small(n // size) + " " + name).strip()); n %= size
    return " ".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# NDA NOT SIGNED — the datasheet stops after the gas composition
# ─────────────────────────────────────────────────────────────────────────────
GAS_WORDS = ("methane", "ethane", "propane", "butane", "pentane", "hexane",
             "nitrogen", "carbon dioxide", "co2", "h2s", "hydrogen", "oxygen",
             "helium", "argon", "water", "ch4", "c2h6", "c3h8", "n2")


def cut_after_gas_composition(src, dst):
    """Keep the datasheet pages up to the gas composition table. Returns
    (path, kept, total) or None when no composition table is found."""
    reader = PdfReader(src)
    texts = []
    for pg in reader.pages:
        try:
            texts.append((pg.extract_text() or "").lower())
        except Exception:
            texts.append("")
    best, top = None, 0
    for i, t in enumerate(texts):
        if "composition" not in t and not re.search(r"mol(?:e|ar)?\s*%", t):
            continue
        s = 5 * ("composition" in t) + sum(1 for w in GAS_WORDS
                                           if re.search(rf"\b{re.escape(w)}\b", t))
        if s > top:
            best, top = i, s
    if best is None:
        return None
    if best == len(texts) - 1:
        return src, len(texts), len(texts)
    w = PdfWriter()
    for pg in reader.pages[:best + 1]:
        w.add_page(pg)
    with open(dst, "wb") as fh:
        w.write(fh)
    return dst, best + 1, len(texts)


# ─────────────────────────────────────────────────────────────────────────────
# "Baker Hughes Confidential" at the foot of every page
# ─────────────────────────────────────────────────────────────────────────────
CONFIDENTIAL = "Baker Hughes Confidential"
CONF_RX = re.compile(r"baker\s*hughes\s*confidential", re.I)


def _mark_font():
    """Calibri as in the Word master, or its metric twin Carlito."""
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        if "YBT-Calibri" in pdfmetrics.getRegisteredFontNames():
            return "YBT-Calibri"
        for p in (os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts",
                               "calibri.ttf"),
                  "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf"):
            if os.path.exists(p):
                pdfmetrics.registerFont(TTFont("YBT-Calibri", p))
                return "YBT-Calibri"
    except Exception:
        pass
    return "Helvetica"


def stamp_confidential(path):
    """Write the mark, centred at the foot of the page as in the Word master,
    on every page that does not carry it yet. Returns the number of pages
    marked."""
    import io
    reader = PdfReader(path)
    try:
        writer = PdfWriter(clone_from=reader)
        pages = writer.pages
    except TypeError:
        writer = PdfWriter()
        for pg in reader.pages:
            writer.add_page(pg)
        pages = writer.pages
    font, n = _mark_font(), 0
    for pg in pages:
        try:
            if CONF_RX.search(pg.extract_text() or ""):
                continue
        except Exception:
            pass
        try:
            if pg.get("/Rotate") not in (None, 0):
                pg.transfer_rotation_to_content()
            box = pg.mediabox
            w, h = float(box.width), float(box.height)
            buf = io.BytesIO()
            c = rl_canvas.Canvas(buf, pagesize=(w, h))
            tw = c.stringWidth(CONFIDENTIAL, font, 11)
            c.setFillColorRGB(1, 1, 1)
            c.rect((w - tw) / 2 - 4, 13, tw + 8, 15, fill=1, stroke=0)
            c.setFillColorRGB(1, 0, 0)
            c.setFont(font, 11)
            c.drawCentredString(w / 2, 17, CONFIDENTIAL)
            c.save()
            mark = PdfReader(buf).pages[0]
            x0, y0 = float(box.left), float(box.bottom)
            if x0 or y0:
                from pypdf import Transformation
                pg.merge_transformed_page(mark, Transformation().translate(x0, y0))
            else:
                pg.merge_page(mark)
            n += 1
        except Exception:
            continue
    tmp = path + ".mark"
    with open(tmp, "wb") as fh:
        writer.write(fh)
    os.replace(tmp, path)
    return n


# ─────────────────────────────────────────────────────────────────────────────
# PROJECT FOLDER — "Opportunity structure\100.Technical Proposal - DTI"
# ─────────────────────────────────────────────────────────────────────────────
SECTION_DIR = re.compile(r"^\s*section\s*0*(\d+)(?!\d)", re.I)
CATEGORY_WORDS = {"DTI": ("dti",), "BUDGET": ("budget",), "FIRM": ("firm",)}


def proposal_root(path, category):
    """The folder holding the Section N folders. Picking the opportunity
    folder above it is enough: the proposal of the offer category is found
    inside."""
    def sections(d):
        return sum(1 for n in _safe_listdir(d)
                   if SECTION_DIR.match(n) and os.path.isdir(os.path.join(d, n)))
    if not path or sections(path) >= 2:
        return path
    best, top = path, 0
    for dp, dns, _f in os.walk(path):
        if dp[len(path):].count(os.sep) >= 2:
            dns[:] = []
        for n in dns:
            d = os.path.join(dp, n)
            s = sections(d)
            if s < 2:
                continue
            low = n.lower()
            s += 5 * ("technical proposal" in low)
            s += 8 * any(w in low for w in CATEGORY_WORDS.get(category, ()))
            if s > top:
                best, top = d, s
    return best


# ─────────────────────────────────────────────────────────────────────────────
# TECHNICAL PROPOSAL — DOCUMENT STRUCTURE# ─────────────────────────────────────────────────────────────────────────────
# TECHNICAL PROPOSAL — DOCUMENT STRUCTURE
#
#   SRC_DB    standard document, taken from the local ICL document database
#   SRC_USER  project-specific document, must be provided in the project folder
#   SRC_TOOL  built by this tool (scope of supply, commercial conditions…)
#
#   keys  : filename keywords used to locate the document
#   dyn   : dynamic keyword built from the current configuration
#           ("cooling" / "casing" / "plc" / "config" / "motor")
# ─────────────────────────────────────────────────────────────────────────────
SRC_DB, SRC_USER, SRC_TOOL = "DB", "USER", "TOOL"

DOC_EXT = (".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".pptx", ".ppt",
           ".dwg", ".dxf", ".png", ".jpg", ".jpeg", ".tif", ".txt", ".md", ".csv")


def _D(code, title, src, keys=(), dyn=None, sub=None, ref=None, tdyn=None,
       note="", db=()):
    """One document row.

    code : unique key          ref : number shown in the interface
    tdyn : title that follows the configuration (e.g. the PLC brand)
    """
    return dict(code=code, title=title, src=src, keys=list(keys), dyn=dyn,
                sub=sub or [], ref=ref if ref is not None else code,
                tdyn=tdyn, note=note, db=list(db),
                nda="datasheet" in _flat(title))


def _S(num, title, folder, items):
    return dict(num=num, title=title, folder=folder, items=items)


def _leaves(sec):
    """(item, depth) for every row of a section, deepest level carries a status."""
    for it in sec["items"]:
        yield it, 0
        for s in it["sub"]:
            yield s, 1


# ── Front matter, common to every offer category ─────────────────────────────
def _front_matter():
    return _S(0, "Proposal Front Matter", "Section 0_Front matter", [
        _D("0.1", "Cover Page",                         SRC_TOOL, ("cover",)),
        _D("0.2", "Confidentiality Notice",             SRC_DB,   ("confidentiality",)),
        _D("0.3", "Table of Contents",                  SRC_TOOL, ("table of contents", "toc")),
        _D("0.4", "Technical Proposal (master file)",   SRC_TOOL, ("technical proposal",)),
    ])


# ── DTI / BUDGET : preliminary technical proposal ────────────────────────────
DTI_TECHNICAL = ["gad", "pid", "plc", "amb", "cables", "gepc", "utility",
                 "safety", "landing", "hookup", "piping", "painting", "packing",
                 "qcp", "dilution", "spare", "tools", "subsup", "trs", "ce",
                 "vdr", "reflist"]

DOC_STRUCT_DTI = [
    _front_matter(),
    _S(1, "Executive Summary & Basic Concepts", "Section 1_Executive Summary", [
        _D("1.0", "Proposal Highlights",                SRC_DB,   ("proposal", "highlight"), ref="—"),
        _D("1.1", "ICL Main Technical Features",        SRC_DB,   ("main", "feature")),
        _D("1.2", "Centrifugal Compressors Description",SRC_DB,   ("centrifugal", "description")),
        _D("1.3", "Cooling System Concept",             SRC_DB,   ("cooling", "concept"), db=("cooling",)),
        _D("1.4", "ATEX Protection",                    SRC_DB,   ("atex",)),
        _D("1.5", "Gas Quality Requirements",           SRC_DB,   ("gas quality", "quality requirement")),
    ]),
    _S(2, "Definition of Supply", "Section 2_Definition of supply", [
        _D("2.1", "Scope of Supply",                    SRC_TOOL, ("scope", "supply"), db=("scope",)),
        _D("2.2", "Preliminary Compressor Datasheet",   SRC_USER, ("datasheet", "data sheet")),
        _D("2.3", "Compressor Performance Curves",      SRC_USER, ("curve", "perfo")),
    ]),
    _S(3, "Technical", "Section 3_Technical", [
        _D(f"3.{i}", DB_BY_ID[d]["title"],
           SRC_USER if DB_BY_ID[d]["rule"] == "project" else SRC_DB,
           (), db=(d,))
        for i, d in enumerate(DTI_TECHNICAL, 1)
    ]),
    _S(11, "Commercial Conditions", "Section 11_Commercial", [
        _D("11", "Commercial Conditions",               SRC_TOOL, ("commercial", "price")),
    ]),
]

# ── FIRM : detailed technical proposal — follows the proposal table of contents
DOC_STRUCT_FIRM = [
    _S(0, "Front Matter", "Section 0_Front matter", [
        _D("0.1", "Cover Page",                          SRC_TOOL, ("cover",)),
        _D("0.2", "Confidentiality Notice",              SRC_DB,   ("confidentiality",)),
        _D("0.3", "Table of Contents",                   SRC_TOOL, ("table of contents", "toc")),
        _D("0.4", "GLOSSARY",                            SRC_DB,   ("glossary",)),
    ]),
    _S(1, "Executive Summary & Basic Concepts", "Section 1_Executive Summary", [
        _D("1.0", "Proposal Highlights",                 SRC_DB,   ("proposal", "highlight"), ref="—"),
        _D("1.1", "ICL Main Technical Features",         SRC_DB,   ("main", "feature")),
        _D("1.2", "Centrifugal Compressors Description", SRC_DB,   ("centrifugal", "description")),
        _D("1.3", "Cooling System Concept",              SRC_DB,   ("cooling", "concept"), dyn="cooling"),
        _D("1.4", "Atex Protection",                     SRC_DB,   ("atex",)),
    ]),
    _S(2, "Definition of Supply", "Section 2_Definition of Supply", [
        _D("2.1", "Scope of Supply",                     SRC_TOOL, ("scope", "supply")),
        _D("2.2", "Vendor List",                         SRC_DB,   ("vendor", "sub-vendor")),
    ]),
    _S(3, "Compressor Performance", "Section 3_Compressor Performance", [
        _D("3.1", "Compressor Datasheets",               SRC_USER, ("datasheet", "data sheet")),
        _D("3.2", "Compressor Curves",                   SRC_USER, ("curve", "perfo")),
    ]),
    _S(4, "ICL Safety Concept", "Section 4_ICL Safety Concept", [
        _D("4", "ICL Safety Concept",                    SRC_DB,   ("safety", "concept")),
    ]),
    _S(5, "Electrical System", "Section 5_Electrical System", [
        _D("5", "Electrical System",                     SRC_USER, ("electrical",)),
    ]),
    _S(6, "Auxiliaries", "Section 6_Auxiliaries", [
        _D("6.1", "AMB Active Magnetic Bearings CONTROL",SRC_DB,   ("amb", "magnetic bearing")),
        _D("6.2", "PID",                                 SRC_USER, ("pid", "p&id"), dyn="cooling"),
        _D("6.3", "Piping Lines Specification (Piping Class)", SRC_DB, ("piping",)),
        _D("6.4", "Hook Up Drawings",                    SRC_DB,   ("hook",)),
    ]),
    _S(7, "General Arrangement", "Section 7_General Arrangement", [
        _D("7.1", "Foundation Drawings",                 SRC_USER, ("foundation",)),
        _D("7.2", "Overall Dimensions of Units and Auxiliaries", SRC_USER,
           ("general arrangement", "dimension"), dyn="casing"),
    ]),
    _S(8, "Test and Quality", "Section 8_Test and quality", [
        _D("8", "Test and Quality",                      SRC_DB,   ("qcp", "quality control", "test")),
    ]),
    _S(9, "Miscellaneous", "Section 9_Miscellaneous", [
        _D("9.1", "Gas Quality Requirement",             SRC_DB,   ("gas quality", "quality requirement")),
        _D("9.2", "Landing Bearings Philosophy (Auxiliary Bearings)", SRC_DB,
           ("landing", "auxiliary bearing")),
        _D("9.3", "Standard Special Tool",               SRC_DB,   ("special tool", "tool"),
           dyn="motor", tdyn="special_tool"),
        _D("9.4", "Painting Procedure",                  SRC_DB,   ("painting",)),
        _D("9.5", "ICL Packing Specification",           SRC_DB,   ("packing",)),
        _D("9.6", "Dynamic Simulation",                  SRC_USER, ("dynamic", "simulation")),
    ]),
    _S(10, "Comment and Exceptions", "Section 10_Comment and Exceptions", [
        _D("10", "Comment and Exceptions",               SRC_USER, ("comment", "exception")),
    ]),
    _S(11, "Technical and Regulation Standards", "Section 11_Technical and Regulation standards", [
        _D("11", "Technical and Regulation Standards",   SRC_DB,   ("standard", "regulation")),
    ]),
    _S(12, "Vendor Document Register (VDR)", "Section 12_Vendor document register", [
        _D("12", "Vendor Document Register (VDR)",       SRC_USER, ("vdr", "document register")),
    ]),
    _S(13, "Control System", "Section 13_Control system", [
        _D("13.1", "Siemens – PLC Configuration",        SRC_DB,   ("plc", "configuration"),
           dyn="plc", tdyn="plc_brand"),
        _D("13.2", "Control System Architecture",        SRC_DB,   ("architecture", "interconnect"), dyn="plc"),
        _D("13.3", "Unit Control Panel Layout",          SRC_DB,   ("ucp", "panel layout")),
        _D("13.4", "Dilution Sequence",                  SRC_DB,   ("dilution",)),
    ]),
    _S(14, "Commissioning & Startup Spare Parts", "Section 14_Commissioning startup spare parts", [
        _D("14", "Commissioning & Startup Spare Parts",  SRC_USER, ("spare", "commissioning")),
    ]),
    _S(15, "Utility and Load List", "Section 15_Utility and load list", [
        _D("15", "Utility and Load List",                SRC_USER, ("utility", "load list")),
    ]),
    _S(16, "ICL Ref List", "Section 16_ICL ref list", [
        _D("16", "ICL Reference List",                   SRC_DB,   ("reference", "ref list")),
    ]),
]

# ── Local standard database layout (created next to the tool) ────────────────
DB_FOLDERS = [
    ("01_Basic_Concepts",    "ICL features, compressor description, ATEX, gas quality"),
    ("02_Cooling_Concepts",  "One document per cooling system concept"),
    ("03_GAD",               "Typical general arrangement drawings, per casing size"),
    ("04_PID",               "Typical P&IDs, per cooling concept"),
    ("05_Control_System",    "Interconnecting diagrams, PLC configuration, UCP layout"),
    ("06_Utility_Load_List", "Typical utility and electrical load lists"),
    ("07_Vendor_References", "Approved sub-vendor list, ICL reference list"),
    ("08_Standards_Quality", "Standards, QCP, test procedures, painting, packing"),
    ("09_Safety",            "ICL safety concept, landing bearing philosophy"),
    ("10_Misc",              "Glossary, hook-up, piping class, special tools"),
]


def _flatten_json(obj, prefix=""):
    out = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            out += _flatten_json(v, f"{prefix}{k} ")
    elif isinstance(obj, list):
        for v in obj:
            out += _flatten_json(v, prefix)
    else:
        out.append((prefix.strip(), obj))
    return out


def _num(txt, default=0.0):
    """Read a number the way people type it: 3 500 000,50 · 3,500,000.50 · 21%."""
    if txt is None:
        return default
    t = str(txt).strip().replace("€", "").replace("%", "")
    t = t.replace("\u00a0", "").replace("\u202f", "").replace(" ", "")
    if "," in t and "." in t:                    # 3,500,000.50 or 3.500.000,50
        t = (t.replace(",", "") if t.rfind(".") > t.rfind(",")
             else t.replace(".", "").replace(",", "."))
    elif "," in t:
        head, _, tail = t.rpartition(",")
        t = (head.replace(",", "") + "." + tail) if len(tail) in (1, 2) \
            else t.replace(",", "")
    try:
        return float(t)
    except ValueError:
        return default


def _tool_dir():
    """Folder the tool runs from (works for a frozen executable too)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()



def _walk_docs(root):
    """Return every document-like file under root."""
    out = []
    if not root or not os.path.isdir(root):
        return out
    for dp, dns, fns in os.walk(root):
        dns[:] = [d for d in dns if not d.startswith(".")]
        for fn in fns:
            if fn.startswith("~$") or fn.startswith("."):
                continue
            if fn.lower().endswith(DOC_EXT):
                out.append(os.path.join(dp, fn))
    return out


def _norm(txt):
    return (txt.lower().replace("_", " ").replace("-", " ")
                       .replace(".", " ").replace("&", " "))


# words too common to identify a document on their own
GENERIC_KEYS = {"icl", "list", "gas", "option", "test", "data"}


def _match_doc(files, keys, dyn_keys=()):
    """Best matching file.

    A file is a candidate when it matches two keywords, one specific keyword,
    or one keyword coming from the current configuration (cooling concept,
    casing size, PLC brand…). Configuration keywords weigh three times more so
    that the right variant wins over a generic file.
    """
    best, best_score = None, 0
    for path in files:
        name = _norm(os.path.basename(path))
        stem = _norm(path)
        score, specific, dyn_hit = 0, 0, False
        for k in keys:
            k = _norm(k).strip()
            if k and k in name:
                score += 1
                if k not in GENERIC_KEYS:
                    specific += 1
        for k in dyn_keys:
            k = _norm(str(k)).strip()
            if len(k) > 2 and (k in name or k in stem):
                score += 3
                dyn_hit = True
        if score < 2 and specific == 0 and not dyn_hit:
            continue
        if score > best_score:
            best, best_score = path, score
    return best, best_score

# ─────────────────────────────────────────────────────────────────────────────
# SCOPE OF SUPPLY
#
# One line = one item, one description, one supplier.
#
#   item     what is supplied, as printed
#   desc     the specification printed under it; text, or a function of the
#            current configuration (cooling, PLC, transformer…)
#   resp     who supplies it by default for (DTI, BUDGET, FIRM):
#            "bh" Baker Hughes · "cust" customer · "opt" option, priced apart
#   ask      the default is only a suggestion: the line stays flagged until
#            somebody picks a supplier
#   auto     a choice already made on the technical page decides the supplier
#   applies  the line only exists for some configurations
#   hint     guidance for the engineer, never printed
# ─────────────────────────────────────────────────────────────────────────────
RESP_NAME = {"bh": "Baker Hughes", "cust": "Customer", "opt": "Option"}
COOLING_TEXT = {"TCV": "Temperature control valve (TCV)", "Fan": "Fan"}


def _build_scope_table():
    B, C, O = "bh", "cust", "opt"

    def R(key, item, desc="", resp=B, section=None, ask=False, hint="",
          auto=None, applies=None):
        if isinstance(resp, str):
            resp = (resp, resp, resp)
        return dict(key=key, item=item, desc=desc,
                    resp=dict(zip(("DTI", "BUDGET", "FIRM"), resp)),
                    section=section, ask=ask, hint=hint, auto=auto,
                    applies=applies)

    def cooling(a):
        f = _concept_features(a.cooling_system.get())
        return COOLING_TEXT["Fan"] if f["fan"] else \
            COOLING_TEXT["TCV"] if f["tcv"] else ""

    def ucp(a):
        opt = a.plc_option.get()
        return a.ucp_type.get() + ("" if opt == "Base Case" else f", {opt.lower()}")

    def feat(a, key):
        return _concept_features(a.cooling_system.get())[key]

    def motor_cooler(a):
        return feat(a, "cooler") and not feat(a, "interstage")

    rows = [
        # ── GENERAL ───────────────────────────────────────────────────────
        R("icl_type", "ICL type", lambda a: a.config_code.get(),
          section="GENERAL"),
        R("mgv_type", "MGV motor type", lambda a: a.motor_type.get()),
        R("vfd", "Variable frequency drive (VFD)",
          hint="GE Power Conversion drive: see the GEPC technical offer."),
        R("transfo", "Transformer and associated instrumentation",
          lambda a: a.transfo_type.get()),
        R("vfd_cooler", "VFD cooler and associated instrumentation",
          lambda a: ("Air-cooled" if a.vfd_cooling.get() == "Cooler"
                     else "Water-cooled chiller")),
        R("thrust", "Thrust bearing on motor",
          applies=lambda a: a.thrust_bearing.get(),
          hint="Follows “Thrust Bearing Included” on the technical page."),
        R("e300_comp", "One (1) E300V2 panel for compressor bearings control"),
        R("e300_drv", "One (1) E300V2 panel for driver bearings control"),
        R("comm_sp", "Set of commissioning spare parts"),
        R("export_pkg", "Export shipment packing",
          "According to the Baker Hughes packing specification"),
        R("anchoring", "Anchoring bolts"),

        # ── COMPRESSOR SKID ───────────────────────────────────────────────
        R("skid", "Compressor skid", section="COMPRESSOR SKID"),
        R("painting", "Painting",
          "According to the Baker Hughes painting specification"),
        R("icl_flt", "ICL cooling line filter", lambda a: a.filter_type.get()),
        R("icl_line", "ICL cooling line", cooling,
          hint="Follows the cooling system concept on the technical page."),
        R("interstage", "Interstage cooler",
          applies=lambda a: feat(a, "interstage"),
          auto=lambda a: B if feat(a, "interstage") else None,
          hint="Listed when the cooling concept has an interstage cooler."),
        R("mot_cooler", "Motor cooling gas cooler",
          applies=motor_cooler, auto=lambda a: B if motor_cooler(a) else None,
          hint="Listed when the cooling concept has a cooler."),
        R("main_jb", "Main motor junction box"),
        R("amb_jb", "AMB junction boxes"),
        R("coupling", "Coupling"),
        R("drain", "Drain lines (ICL and filter)",
          "Terminal points at skid edge, as per P&ID"),
        R("vent", "Vent line"),
        R("cable_pip", "Cables and piping within skid limit"),
        R("therm_ins", "Thermal insulation and personnel protection on "
                       "equipment and pipes", "As per P&ID"),
        R("weather", "Weather protection",
          lambda a: "Compressor enclosure" if a.comp_enclosure.get() else "",
          ask=True, auto=lambda a: B if a.comp_enclosure.get() else None,
          hint="Baker Hughes when “Compressor Enclosure” is ticked on the "
               "technical page."),
        R("instr_skid", "Instrumentation within skid limit"),
        R("instr_panel", "Instrumentation panel", resp=(C, B, B), ask=True),

        # ── ELECTRICAL ITEMS ──────────────────────────────────────────────
        R("cable_gl", "Cable glands for Baker Hughes cable connections",
          section="ELECTRICAL ITEMS"),
        R("cable_gl_c", "Cable glands for customer cable connections", resp=C),
        R("ucp", "Unit control panel (UCP)", ucp),
        R("dilution", "Dilution sequence before start-up",
          "Integrated in the UCP logic",
          hint="See the dilution sequence document."),
        R("pwr_cables", "Power cables between VFD and motor", ask=True,
          hint="Depends on the cable lengths on site: see the cables document."),
        R("mcc", "Motor control centre (MCC)", resp=(C, B, B), ask=True,
          auto=lambda a: B if a.mcc.get() else None,
          hint="Baker Hughes when “MCC” is ticked on the technical page."),
        R("ups", "UPS", resp=(C, B, B), ask=True),
        R("elec_room", "Electrical room (E-house)", resp=(C, B, B), ask=True,
          auto=lambda a: B if a.ehouse.get() else None,
          hint="Baker Hughes when “E-House” is ticked on the technical page."),
        R("harmonic", "Harmonic filter", resp=(C, C, B), ask=True),

        # ── SERVICE ───────────────────────────────────────────────────────
        R("hazop", "HAZOP and SIL review",
          "Two days for two specialists. Design modifications following the "
          "review are handled as change orders",
          section="SERVICE"),
        R("factory_t", "Factory tests, inspection and certification",
          "According to the quality control plan (QCP)"),
        R("test_type", "Tests",
          lambda a: ("MRT, performance test BEC PTC type 2, full load type 1"
                     if a.num_phases.get() == "2" else
                     "MRT, performance test Boitard, full load type 1")),
        R("torsional", "Torsional analysis of complete train"),
        R("lateral", "Lateral analysis of compressor rotor and motor rotor",
          lambda a: ("Two rotors with flexible coupling: the analysis is done "
                     "separately for driver and compressor"
                     if a.back_to_back.get() == "Back to Back" else "")),
        R("atex", "ATEX inspection and certification",
          "No additional certification provided"),
        R("third_p", "Third-party inspections and approvals",
          "Including the related costs", ask=True),
        R("manual", "Instruction manual for maintenance and operation",
          "Documentation and warning signs in English"),
        R("certif", "Certification, calculation and job documentation",
          ask=True),
        R("vh_study", "Voltage harmonic study", ask=True),
        R("vfd_fill", "VFD cooling system mixture, first fill", ask=True),
        R("models_3d", "3D models of the compressor skid", ask=True),
        R("supervision", "Supervision of installation and commissioning",
          resp=O, ask=True, hint="Usually quoted per man-day."),
        R("training", "Training of operation and maintenance staff",
          resp=O, ask=True),
        R("incoterm_r", "Delivery terms",
          lambda a: f"{a.incoterm.get()} (Incoterms 2020)"),
        R("storage", "Equipment storage", "If required", ask=True),

        # ── LOOSE ITEMS ───────────────────────────────────────────────────
        R("spec_tools", "Special tools for compressor maintenance",
          "Planned maintenance only", section="LOOSE ITEMS", ask=True),
        R("asv_ctrl", "ASV control instrumentation",
          "For ASV control only. Standard: orifice plate"),
        R("antisurge", "Process gas antisurge valves", ask=True,
          auto=lambda a: B if a.antisurge_valve.get() else None,
          hint="Baker Hughes when “Antisurge Valve” is ticked on the "
               "technical page."),
        R("hot_byp", "Hot bypass valve",
          "If required: to be confirmed by dynamic simulation during detail "
          "engineering. Supplied loose", resp=O, ask=True),
        R("temp_flt", "Compressor temporary filter (temporary strainer)",
          "For commissioning and first start-up only. Conical filter as per "
          "Baker Hughes standard, sized on the suction flange"),
        R("spool_tmp", "Spool piece for temporary filter", ask=True),
        R("instr_tmp", "Instrumentation of temporary filter",
          "Supplied loose (one PDIT)"),
        R("ctr_flg", "Counter flanges (compressor, ASV…)", ask=True),
        R("lifting", "Skid lifting device", ask=True),
        R("throttle", "Throttling valve", "If required", ask=True),
        R("gas_valves", "Process gas valves and supports", ask=True),
        R("field_ins", "Field process gas instrumentation", ask=True),

        # ── SPARE PARTS ───────────────────────────────────────────────────
        R("cap_sp", "Capital spare parts", resp=O, ask=True,
          section="SPARE PARTS"),
        R("op_sp_2y", "Two-year operational spare parts", resp=O, ask=True),

        # ── SITE AND INSTALLATION ─────────────────────────────────────────
        R("foundations", "Foundations and civil works", resp=C,
          section="SITE AND INSTALLATION"),
        R("erection", "Installation and erection on site", resp=C),
        R("utilities", "Utilities at battery limits",
          "Power supply, instrument air, nitrogen", resp=C),
        R("gas_det", "Site gas detection", resp=C),
        R("proc_pip", "Process piping", resp=C),
        R("interconn", "Interconnection between Baker Hughes equipment",
          resp=C, hint="Turnkey job: see the dedicated table."),
        R("acoustic", "Acoustic insulation", ask=True),
        R("walkways", "Walkways and ladders",
          "Not needed within the Baker Hughes skid", resp=C),
        R("air_comp", "Instrument air compressor", resp=C),
        R("amb_cable", "AMB cables", ask=True,
          hint="Not to be included systematically."),
    ]
    section = ""
    for r in rows:                  # every line knows its section
        section = r["section"] or section
        r["section"] = section
    return rows


SCOPE_TABLE = _build_scope_table()

FIRM_LIBRARY = {"1.3": ["cooling"], "2.1": ["scope"], "2.2": ["subsup"],
                "4": ["safety"], "5": ["gepc", "cables"], "6.1": ["amb"],
                "6.2": ["pid"], "6.3": ["piping"], "6.4": ["hookup"],
                "7.2": ["gad"], "8": ["qcp"], "9.2": ["landing"],
                "9.3": ["tools"], "9.4": ["painting"], "9.5": ["packing"],
                "10": ["ce"], "11": ["trs"], "12": ["vdr"], "13.1": ["plc"],
                "13.4": ["dilution"], "14": ["spare"], "15": ["utility"],
                "16": ["reflist"]}
for _sec in DOC_STRUCT_FIRM:
    for _it in _sec["items"]:
        for _x in [_it] + _it["sub"]:
            if _x["code"] in FIRM_LIBRARY:
                _x["db"] = list(FIRM_LIBRARY[_x["code"]])


# ═════════════════════════════════════════════════════════════════════════════
# INCOTERMS 2020 — who pays what, where the risk passes
#
# The logistics chain is split into eight steps. For each rule, `seller` says
# which steps the seller pays for, `risk` is the last step still at the
# seller's risk, and `insurance` says whether cargo insurance is on the seller.
# ═════════════════════════════════════════════════════════════════════════════
INCOTERM_STAGES = [
    ("Packing",      "Export packing at the works"),
    ("Loading",      "Loading at the seller's premises"),
    ("Inland O",     "Inland carriage, origin, and origin terminal charges"),
    ("Export cust.", "Export customs clearance"),
    ("Main carriage","Main carriage: sea, air, rail or road freight"),
    ("Import cust.", "Import customs clearance, duties and taxes"),
    ("Inland D",     "Inland carriage to the named place of destination"),
    ("Unloading",    "Unloading at the named place of destination"),
]

INCOTERM_DATA = {
    "EXW": dict(name="Ex Works", mode="Any transport mode", insurance=False,
                seller=[1, 0, 0, 0, 0, 0, 0, 0], risk=0,
                note="The seller only makes the packed goods available at its own "
                     "premises. Everything after that, including loading and "
                     "export clearance, is on the buyer."),
    "FCA": dict(name="Free Carrier", mode="Any transport mode", insurance=False,
                seller=[1, 1, 1, 1, 0, 0, 0, 0], risk=3,
                note="The seller clears the goods for export and hands them to the "
                     "carrier named by the buyer. Risk passes at that handover."),
    "FAS": dict(name="Free Alongside Ship", mode="Sea and inland waterway only",
                insurance=False, seller=[1, 1, 1, 1, 0, 0, 0, 0], risk=3,
                note="The seller delivers alongside the vessel at the port of "
                     "shipment, export cleared. Loading on board is on the buyer."),
    "FOB": dict(name="Free On Board", mode="Sea and inland waterway only",
                insurance=False, seller=[1, 1, 1, 1, 0, 0, 0, 0], risk=3,
                note="The seller delivers on board the vessel. Risk passes once the "
                     "goods are on board; freight is on the buyer."),
    "CFR": dict(name="Cost and Freight", mode="Sea and inland waterway only",
                insurance=False, seller=[1, 1, 1, 1, 1, 0, 0, 0], risk=3,
                note="The seller pays the main carriage but risk still passes when "
                     "the goods are on board at origin. Cost and risk split."),
    "CIF": dict(name="Cost, Insurance and Freight",
                mode="Sea and inland waterway only", insurance=True,
                seller=[1, 1, 1, 1, 1, 0, 0, 0], risk=3,
                note="As CFR, plus minimum cargo insurance taken out by the seller "
                     "for the benefit of the buyer."),
    "CPT": dict(name="Carriage Paid To", mode="Any transport mode", insurance=False,
                seller=[1, 1, 1, 1, 1, 0, 0, 0], risk=2,
                note="The seller pays carriage to the named destination, but risk "
                     "passes as soon as the goods are handed to the first carrier."),
    "CIP": dict(name="Carriage and Insurance Paid To", mode="Any transport mode",
                insurance=True, seller=[1, 1, 1, 1, 1, 0, 0, 0], risk=2,
                note="As CPT, plus all-risks cargo insurance taken out by the "
                     "seller for the benefit of the buyer."),
    "DAP": dict(name="Delivered At Place", mode="Any transport mode", insurance=False,
                seller=[1, 1, 1, 1, 1, 0, 1, 0], risk=6,
                note="The seller delivers to the named place, ready for unloading. "
                     "Import clearance and unloading are on the buyer."),
    "DPU": dict(name="Delivered at Place Unloaded", mode="Any transport mode",
                insurance=False, seller=[1, 1, 1, 1, 1, 0, 1, 1], risk=7,
                note="The only rule where the seller unloads at destination. Import "
                     "clearance stays on the buyer."),
    "DDP": dict(name="Delivered Duty Paid", mode="Any transport mode", insurance=False,
                seller=[1, 1, 1, 1, 1, 1, 1, 0], risk=6,
                note="Maximum obligation for the seller: delivered at destination, "
                     "import duties and taxes paid. Unloading is on the buyer."),
}


def incoterm_spec(code):
    sp = INCOTERM_DATA.get(code)
    if not sp:
        return None
    sp = dict(sp)
    sp["code"] = code
    sp["cost"] = max(i for i, v in enumerate(sp["seller"]) if v)
    sp["risk_flags"] = [1 if i <= sp["risk"] else 0 for i in range(len(INCOTERM_STAGES))]
    return sp


def draw_incoterm_tk(cv, code, w, h):
    """Draw the selected rule on a tkinter canvas."""
    cv.delete("all")
    sp = incoterm_spec(code)
    if not sp:
        return
    n = len(INCOTERM_STAGES)
    x0, pad = 74, 12
    cw = max(38, (w - x0 - pad) / n)

    cv.create_text(pad, 14, anchor="w", text=f"{code} — {sp['name']}",
                   font=("Segoe UI", 11, "bold"), fill=BH_INK)
    cv.create_text(w - pad, 14, anchor="e", text=sp["mode"],
                   font=FONT_SM, fill=BH_GRAY)

    for i, (short, _long) in enumerate(INCOTERM_STAGES):
        cv.create_text(x0 + i*cw + cw/2, 36, text=short, font=("Segoe UI", 7),
                       fill=BH_GRAY, width=cw - 4)

    for li, (label, flags) in enumerate([("Costs", sp["seller"]),
                                         ("Risk", sp["risk_flags"])]):
        y = 48 + li*32
        cv.create_text(x0 - 10, y + 12, anchor="e", text=label,
                       font=("Segoe UI", 8, "bold"), fill=BH_INK)
        for i in range(n):
            x = x0 + i*cw
            cv.create_rectangle(x + 1, y, x + cw - 1, y + 24,
                                fill=BH_OK if flags[i] else BH_GRID, outline="")
        cv.create_text(x0 + 4, y + 12, anchor="w", text="SELLER",
                       font=("Segoe UI", 7, "bold"), fill=BH_WHITE)
        last = flags.index(0) if 0 in flags else n
        if last < n:
            cv.create_text(x0 + last*cw + 4, y + 12, anchor="w", text="BUYER",
                           font=("Segoe UI", 7, "bold"), fill=BH_GRAY)

    # transfer markers
    ymk = 48 + 2*32 + 4
    for idx, txt, col in [(sp["cost"], "costs pass here", BH_OK),
                          (sp["risk"], "risk passes here", BH_WARN)]:
        x = x0 + (idx + 1)*cw
        cv.create_polygon(x - 5, ymk + 10, x + 5, ymk + 10, x, ymk,
                          fill=col, outline="")
        cv.create_text(x + 8, ymk + 5, anchor="w", text=txt,
                       font=("Segoe UI", 7, "bold"), fill=col)
        ymk += 15

    ins = "Cargo insurance paid by the seller" if sp["insurance"] else \
          "No insurance obligation on the seller"
    cv.create_text(pad, h - 10, anchor="w", text=ins, font=FONT_SM,
                   fill=BH_OK if sp["insurance"] else BH_GRAY)


def draw_incoterm_matrix_tk(cv, selected, w, h):
    """Compare every rule; the selected one is highlighted."""
    cv.delete("all")
    n = len(INCOTERM_STAGES)
    x0, pad, rh = 74, 12, 17
    cw = max(34, (w - x0 - pad - 58) / n)
    for i, (short, _l) in enumerate(INCOTERM_STAGES):
        cv.create_text(x0 + i*cw + cw/2, 10, text=short, font=("Segoe UI", 7),
                       fill=BH_GRAY, width=cw - 4)
    cv.create_text(x0 + n*cw + 26, 10, text="Insurance", font=("Segoe UI", 7),
                   fill=BH_GRAY)
    y = 22
    for code in INCOTERMS:
        sp = incoterm_spec(code)
        if not sp:
            continue
        on = (code == selected)
        if on:
            cv.create_rectangle(pad, y, w - pad, y + rh - 2, fill=BH_SEC, outline="")
        cv.create_text(x0 - 10, y + rh/2 - 1, anchor="e", text=code,
                       font=("Segoe UI", 8, "bold" if on else "normal"),
                       fill=BH_GREEN if on else BH_INK)
        for i in range(n):
            x = x0 + i*cw + cw/2
            if sp["seller"][i]:
                r = 4
                cv.create_oval(x - r, y + rh/2 - 1 - r, x + r, y + rh/2 - 1 + r,
                               fill=BH_OK if on else BH_OK_PALE, outline="")
            else:
                cv.create_line(x - 3, y + rh/2 - 1, x + 3, y + rh/2 - 1,
                               fill=BH_GRID)
        xi = x0 + n*cw + 26
        if sp["insurance"]:
            cv.create_oval(xi - 4, y + rh/2 - 5, xi + 4, y + rh/2 + 3,
                           fill=BH_OK if on else BH_OK_PALE, outline="")
        else:
            cv.create_line(xi - 3, y + rh/2 - 1, xi + 3, y + rh/2 - 1,
                           fill=BH_GRID)
        if on:
            cv.create_text(w - pad - 4, y + rh/2 - 1, anchor="e",
                           text="selected", font=("Segoe UI", 7, "bold"),
                           fill=BH_GREEN)
        y += rh
    cv.create_text(pad, y + 8, anchor="w",
                   text="●  paid by Baker Hughes (seller)        "
                        "–  paid by the customer (buyer)",
                   font=FONT_SM, fill=BH_GRAY)


# ═════════════════════════════════════════════════════════════════════════════
# TECHNICAL PROPOSAL PDF ASSEMBLY
#
# Builds the complete proposal: branded cover, confidentiality notice, table of
# contents with real page numbers, one divider per section, the documents
# themselves (from the project folder or the standard database), generated
# pages for what the tool owns (scope of supply, commercial conditions) and a
# placeholder wherever a document is still missing. Every page is then stamped
# with the running header and footer.
# ═════════════════════════════════════════════════════════════════════════════
try:                                     # merging engine
    from pypdf import PdfReader, PdfWriter
    HAS_PYPDF = True
except ImportError:                      # pragma: no cover
    try:
        from PyPDF2 import PdfReader, PdfWriter
        HAS_PYPDF = True
    except ImportError:
        HAS_PYPDF = False

try:                                     # page generation engine
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.utils import ImageReader
    HAS_REPORTLAB = True
except ImportError:                      # pragma: no cover
    HAS_REPORTLAB = False

PDF_IMG_EXT = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")


def _rl_colour(hex_colour):
    hex_colour = hex_colour.lstrip("#")
    return tuple(int(hex_colour[i:i+2], 16) / 255 for i in (0, 2, 4))


class ProposalPDF:
    """Assembles the offer package into a single PDF."""

    def __init__(self, data, struct, status, options=None, logo=None):
        self.d = data
        self.struct = struct
        self.status = status
        self.o = dict(placeholders=True, stamp=True, scope=True,
                      commercial=True, skip_empty=False)
        self.o.update(options or {})
        self.logo = logo if (logo and os.path.exists(logo)) else None
        self.W, self.H = A4
        self.green = _rl_colour(BH_GREEN)
        self.ink = _rl_colour(BH_INK)
        self.grey = _rl_colour(BH_MUTED)
        self.report = []          # (ref, title, outcome)
        self.section_starts = []  # (label, first page index)

    # ── small drawing helpers ────────────────────────────────────────────
    def _running_head(self, c):
        c.setFont("Helvetica", 8)
        c.setFillColorRGB(*self.ink)
        c.drawString(20*mm, self.H - 15*mm,
                     f"Customer name : {self.d.get('client_name') or '—'}")
        c.drawString(20*mm, self.H - 20*mm,
                     f"{self.d.get('application') or ''} "
                     f"{('— ' + self.d['client_loc']) if self.d.get('client_loc') else ''}".strip())
        if self.logo:
            try:
                c.drawImage(ImageReader(self.logo), self.W - 70*mm, self.H - 24*mm,
                            width=50*mm, height=12*mm, mask="auto",
                            preserveAspectRatio=True, anchor="e")
                return
            except Exception:
                pass
        c.setFont("Helvetica-Bold", 13)
        c.setFillColorRGB(*self.green)
        c.drawRightString(self.W - 20*mm, self.H - 19*mm, "Baker Hughes")

    def _running_foot(self, c, page=None, total=None):
        c.setFont("Helvetica", 7)
        c.setFillColorRGB(*self.grey)
        ref = self.d.get("opp_id") or "—"
        c.drawString(20*mm, 15*mm, f"Baker Hughes Ref: {ref}")
        c.drawString(20*mm, 11*mm,
                     f"Copyright {datetime.now().year} Baker Hughes Company LLC. "
                     "All rights reserved.")
        if page is not None:
            c.drawRightString(self.W - 20*mm, 11*mm,
                              f"Page {page} of {total}" if total else f"Page {page}")

    def _title_block(self, c, title, sub=""):
        c.setFillColorRGB(*self.green)
        c.rect(20*mm, self.H - 42*mm, 3*mm, 12*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 17)
        c.setFillColorRGB(*self.ink)
        c.drawString(26*mm, self.H - 38*mm, title)
        if sub:
            c.setFont("Helvetica", 9)
            c.setFillColorRGB(*self.grey)
            c.drawString(26*mm, self.H - 44*mm, sub)

    def _fit(self, c, text, width, font=("Helvetica", 8)):
        """Truncate to the available width, with an ellipsis."""
        if c.stringWidth(text, *font) <= width:
            return text
        while text and c.stringWidth(text + "…", *font) > width:
            text = text[:-1]
        return text + "…"

    def _wrap(self, c, text, x, y, width, leading=5.0*mm, font=("Helvetica", 9)):
        c.setFont(*font)
        words, line = text.split(), ""
        for w in words:
            probe = f"{line} {w}".strip()
            if c.stringWidth(probe, *font) <= width:
                line = probe
            else:
                c.drawString(x, y, line); y -= leading; line = w
        if line:
            c.drawString(x, y, line); y -= leading
        return y

    # ── generated pages ──────────────────────────────────────────────────
    def _cover(self, c):
        cat = self.d.get("offer_category", "DTI")
        self._running_head(c)
        c.setFillColorRGB(*self.green)
        c.rect(0, self.H - 120*mm, self.W, 1.2*mm, fill=1, stroke=0)
        c.setFont("Helvetica", 26)
        c.setFillColorRGB(*self.ink)
        if cat in ("DTI", "BUDGET"):
            c.drawCentredString(self.W/2, self.H - 145*mm, "PRELIMINARY")
        c.setFont("Helvetica-Bold", 30)
        c.drawCentredString(self.W/2, self.H - 162*mm, "Technical Proposal")
        c.setFillColorRGB(*self.green)
        c.rect(20*mm, self.H - 172*mm, self.W - 40*mm, 0.8*mm, fill=1, stroke=0)

        y = self.H - 186*mm
        rows = [("Project:", self.d.get("application", "")),
                ("Customer:", self.d.get("client_name", "")),
                ("Location:", self.d.get("client_loc", "")),
                ("Proposal Reference Number:", self.d.get("opp_id", "")),
                ("Revision:", self.d.get("revision", "")),
                ("Configuration:", self.d.get("config_code", "")),
                ("Date:", self.d.get("project_date")
                 or datetime.now().strftime("%d/%m/%Y"))]
        for k, v in rows:
            if not v:
                continue
            c.setFont("Helvetica-Bold", 10); c.setFillColorRGB(*self.ink)
            c.drawString(20*mm, y, k)
            c.setFont("Helvetica", 10)
            c.drawString(85*mm, y, str(v))
            y -= 7*mm

        y -= 4*mm
        c.setFont("Helvetica-Bold", 10)
        c.drawString(20*mm, y, "Contact Persons:"); y -= 8*mm
        for name, role, mail, phone in [
                (self.d.get("sales_name"), "Sales Manager",
                 self.d.get("sales_email"), self.d.get("sales_phone")),
                (self.d.get("adv_name"), "Sales Advisor",
                 self.d.get("adv_email"), self.d.get("adv_phone")),
                (self.d.get("ae_name"), "Application Engineer",
                 self.d.get("ae_email"), self.d.get("ae_phone"))]:
            if not name:
                continue
            c.setFont("Helvetica", 10); c.drawString(20*mm, y, name)
            c.setFont("Helvetica-Bold", 9); c.drawString(20*mm, y - 5*mm, role)
            c.setFont("Helvetica", 9)
            c.setFillColorRGB(*self.grey)
            c.drawString(20*mm, y - 10*mm, " · ".join(x for x in (mail, phone) if x))
            c.setFillColorRGB(*self.ink)
            y -= 18*mm
        self._running_foot(c)
        c.showPage()

    def _confidentiality(self, c):
        self._running_head(c)
        self._title_block(c, "Confidentiality notice")
        txt = ("This document and all information in it are confidential and the "
               "proprietary property of Baker Hughes. It is furnished to, and "
               "accepted by, recipient solely for the use by recipient to evaluate "
               "the Baker Hughes proposal in this document. The contents of the "
               "proposal must not be disclosed to any third party outside of the "
               "recipient organization without the prior written consent of Baker "
               "Hughes. Copying by any means or further distribution of the "
               "proposal, or any of its contents, is strictly prohibited without "
               "Baker Hughes prior written consent, except that reasonably required "
               "reproduction is authorized in connection with recipient evaluation "
               "of the proposal. Recipient shall destroy or return the proposal, and "
               "any copies thereof, on the earlier of request by Baker Hughes or "
               "completion of the proposal evaluation.")
        c.setFillColorRGB(*self.ink)
        self._wrap(c, txt, 20*mm, self.H - 58*mm, self.W - 40*mm)
        self._running_foot(c)
        c.showPage()

    def _toc(self, c, entries, pages):
        """entries: (depth, ref, title). pages: same length, page number or None."""
        per_page = 34
        for chunk in range(0, max(1, len(entries)), per_page):
            self._running_head(c)
            self._title_block(c, "Table of contents" if chunk == 0 else
                              "Table of contents (continued)")
            y = self.H - 58*mm
            for i in range(chunk, min(chunk + per_page, len(entries))):
                depth, ref, title = entries[i]
                pg = pages[i] if pages else None
                x = 20*mm + depth * 8*mm
                if depth == 0:
                    c.setFont("Helvetica-Bold", 9.5); c.setFillColorRGB(*self.ink)
                else:
                    c.setFont("Helvetica", 9); c.setFillColorRGB(*self.ink)
                label = f"{ref}   {title}" if ref and ref != "—" else title
                c.drawString(x, y, self._fit(c, label, 140*mm - depth*8*mm,
                                             ("Helvetica-Bold", 9.5) if depth == 0
                                             else ("Helvetica", 9)))
                if pg:
                    c.setFont("Helvetica", 8.5); c.setFillColorRGB(*self.grey)
                    c.drawRightString(self.W - 20*mm, y, str(pg))
                y -= 5.4*mm
            self._running_foot(c)
            c.showPage()

    def _divider(self, c, sec):
        self._running_head(c)
        c.setFillColorRGB(*self.green)
        c.rect(20*mm, self.H/2 + 6*mm, 25*mm, 1.6*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 26)
        c.setFillColorRGB(*self.ink)
        c.drawString(20*mm, self.H/2 - 8*mm, f"Section {sec.get('disp', sec['num'])}")
        c.setFont("Helvetica-Bold", 19)
        c.setFillColorRGB(*self.green)
        c.drawString(20*mm, self.H/2 - 22*mm, sec["title"])
        self._running_foot(c)
        c.showPage()

    def _placeholder(self, c, ref, title, reason):
        self._running_head(c)
        self._title_block(c, f"{ref}  {title}" if ref and ref != "—" else title)
        c.setFillColorRGB(*_rl_colour(BH_WARN))
        c.rect(20*mm, self.H - 72*mm, self.W - 40*mm, 14*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 11)
        c.setFillColorRGB(1, 1, 1)
        c.drawString(25*mm, self.H - 67*mm, "DOCUMENT TO BE PROVIDED")
        c.setFillColorRGB(*self.ink)
        self._wrap(c, reason, 20*mm, self.H - 84*mm, self.W - 40*mm)
        self._running_foot(c)
        c.showPage()

    def _attachment_note(self, c, ref, title, path):
        self._running_head(c)
        self._title_block(c, f"{ref}  {title}" if ref and ref != "—" else title)
        c.setFillColorRGB(*_rl_colour(BH_TOOL))
        c.rect(20*mm, self.H - 72*mm, self.W - 40*mm, 14*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 11); c.setFillColorRGB(1, 1, 1)
        c.drawString(25*mm, self.H - 67*mm, "SUPPLIED AS A SEPARATE FILE")
        c.setFillColorRGB(*self.ink)
        self._wrap(c, f"This document is attached to the proposal package as "
                      f"{os.path.basename(path)}. Convert it to PDF to have it "
                      f"merged into this file.", 20*mm, self.H - 84*mm,
                   self.W - 40*mm)
        self._running_foot(c)
        c.showPage()

    def _image_page(self, c, ref, title, path):
        self._running_head(c)
        self._title_block(c, f"{ref}  {title}" if ref and ref != "—" else title)
        try:
            img = ImageReader(path)
            iw, ih = img.getSize()
            box_w, box_h = self.W - 40*mm, self.H - 90*mm
            k = min(box_w/iw, box_h/ih)
            c.drawImage(img, (self.W - iw*k)/2, 25*mm, width=iw*k, height=ih*k,
                        preserveAspectRatio=True, mask="auto")
        except Exception:
            c.setFont("Helvetica", 9)
            c.drawString(20*mm, self.H/2, f"Image could not be embedded: {path}")
        self._running_foot(c)
        c.showPage()

    def _scope_pages(self, c):
        rows = self.d.get("scope_of_supply") or []
        if not rows:
            self._placeholder(c, "2.1", "Scope of Supply",
                              "No scope line has been selected in the tool.")
            return 1
        pages, i = 0, 0
        while i < len(rows):
            self._running_head(c)
            self._title_block(c, "Scope of Supply",
                              f"{self.d.get('config_code','')}   "
                              f"{self.d.get('application','')}")
            y = self.H - 60*mm
            c.setFillColorRGB(*self.green)
            c.rect(20*mm, y, self.W - 40*mm, 7*mm, fill=1, stroke=0)
            c.setFont("Helvetica-Bold", 8); c.setFillColorRGB(1, 1, 1)
            c.drawString(22*mm, y + 2*mm, "ITEM")
            c.drawCentredString(132*mm, y + 2*mm, "BH")
            c.drawCentredString(146*mm, y + 2*mm, "CUSTOMER")
            c.drawCentredString(161*mm, y + 2*mm, "OPTION")
            y -= 7*mm
            cur = None
            while i < len(rows) and y > 28*mm:
                r = rows[i]
                if r.get("section") and r["section"] != cur:
                    cur = r["section"]
                    c.setFillColorRGB(*_rl_colour(BH_BG))
                    c.rect(20*mm, y - 1*mm, self.W - 40*mm, 6*mm, fill=1, stroke=0)
                    c.setFont("Helvetica-Bold", 8); c.setFillColorRGB(*self.ink)
                    c.drawString(22*mm, y + 0.8*mm, str(cur))
                    y -= 6.5*mm
                    if y < 28*mm:
                        break
                item = r.get("item") or r.get("label", "")
                desc = r.get("desc", "")
                c.setFont("Helvetica", 8); c.setFillColorRGB(*self.ink)
                shown = self._fit(c, item, 100*mm)
                c.drawString(24*mm, y, shown)
                room = 100*mm - c.stringWidth(shown, "Helvetica", 8) - 3*mm
                if desc and room > 14*mm:
                    c.setFont("Helvetica", 7); c.setFillColorRGB(*self.grey)
                    c.drawString(24*mm + 100*mm - room, y,
                                 self._fit(c, desc, room, ("Helvetica", 7)))
                c.setFont("Helvetica-Bold", 8); c.setFillColorRGB(*self.ink)
                resp = r.get("resp") or ("bh" if r.get("bh") else
                                         "cust" if r.get("customer") else "")
                x = {"bh": 132*mm, "cust": 146*mm, "opt": 161*mm}.get(resp)
                if x:
                    c.drawCentredString(x, y, "X")
                y -= 5.2*mm
                i += 1
            self._running_foot(c)
            c.showPage(); pages += 1
        return pages

    def _commercial_page(self, c):
        """Commercial conditions + the delivery-terms illustration. Two pages."""
        self._running_head(c)
        self._title_block(c, "Commercial Conditions")
        y = self.H - 58*mm
        rows = [("Final price", self.d.get("final_price")),
                ("INCOTERM", self.d.get("incoterm")),
                ("Lead time", self.d.get("lead_time")),
                ("Validity", "90 days from the date of this proposal"),
                ("Proposal reference", self.d.get("opp_id")),
                ("Revision", self.d.get("revision"))]
        for k, v in rows:
            if not v:
                continue
            c.setFont("Helvetica-Bold", 9.5); c.setFillColorRGB(*self.ink)
            c.drawString(20*mm, y, k)
            c.setFont("Helvetica", 9.5)
            c.drawString(70*mm, y, str(v))
            y -= 6.5*mm

        note = incoterm_note(self.d.get("incoterm", ""))
        if note:
            c.setFont("Helvetica-Oblique", 8.5); c.setFillColorRGB(*self.grey)
            y = self._wrap(c, note, 20*mm, y - 2*mm, self.W - 40*mm, leading=4.4*mm)
        y -= 6*mm
        y = self._incoterm_block(c, self.d.get("incoterm", ""), y)

        if self.d.get("clarifications"):
            y -= 8*mm
            c.setFont("Helvetica-Bold", 9.5); c.setFillColorRGB(*self.ink)
            c.drawString(20*mm, y, "Clarifications & notes")
            y -= 6*mm
            for line in self.d["clarifications"].splitlines():
                if y < 30*mm:
                    break
                y = self._wrap(c, line, 20*mm, y, self.W - 40*mm, leading=4.6*mm)
        self._running_foot(c)
        c.showPage()

        self._incoterm_matrix_page(c, self.d.get("incoterm", ""))
        return 2

    # ── delivery terms, drawn ─────────────────────────────────────────────
    def _incoterm_block(self, c, code, y):
        sp = incoterm_spec(code)
        if not sp:
            return y
        n = len(INCOTERM_STAGES)
        x0 = 44*mm
        cw = (self.W - 20*mm - x0) / n

        c.setFillColorRGB(*self.green)
        c.rect(20*mm, y - 1*mm, 3*mm, 6*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 11); c.setFillColorRGB(*self.ink)
        c.drawString(26*mm, y, f"{code} — {sp['name']}")
        c.setFont("Helvetica", 8); c.setFillColorRGB(*self.grey)
        c.drawRightString(self.W - 20*mm, y, sp["mode"])
        y -= 9*mm

        c.setFont("Helvetica", 6)
        for i, (short, _l) in enumerate(INCOTERM_STAGES):
            c.setFillColorRGB(*self.grey)
            c.drawCentredString(x0 + i*cw + cw/2, y, short)
        y -= 3*mm

        for label, flags in [("Costs", sp["seller"]), ("Risk", sp["risk_flags"])]:
            c.setFont("Helvetica-Bold", 8); c.setFillColorRGB(*self.ink)
            c.drawRightString(x0 - 3*mm, y - 5*mm, label)
            for i in range(n):
                if flags[i]:
                    c.setFillColorRGB(*_rl_colour(BH_OK))
                else:
                    c.setFillColorRGB(*_rl_colour(BH_GRID))
                c.rect(x0 + i*cw + 0.3*mm, y - 7*mm, cw - 0.6*mm, 6*mm,
                       fill=1, stroke=0)
            c.setFont("Helvetica-Bold", 6); c.setFillColorRGB(1, 1, 1)
            c.drawString(x0 + 1.5*mm, y - 5*mm, "SELLER")
            last = flags.index(0) if 0 in flags else n
            if last < n:
                c.setFillColorRGB(*_rl_colour(BH_GRAY))
                c.drawString(x0 + last*cw + 1.5*mm, y - 5*mm, "BUYER")
            y -= 9*mm

        for idx, txt, col in [(sp["cost"], "costs pass here", BH_OK),
                              (sp["risk"], "risk passes here", BH_WARN)]:
            x = x0 + (idx + 1)*cw
            c.setFillColorRGB(*_rl_colour(col))
            c.setFont("Helvetica-Bold", 6.5)
            path = c.beginPath()
            path.moveTo(x - 1.4*mm, y - 1*mm); path.lineTo(x + 1.4*mm, y - 1*mm)
            path.lineTo(x, y + 1.4*mm); path.close()
            c.drawPath(path, fill=1, stroke=0)
            c.drawString(x + 2*mm, y - 1*mm, txt)
            y -= 5*mm

        c.setFont("Helvetica-Oblique", 8)
        c.setFillColorRGB(*(_rl_colour(BH_OK) if sp["insurance"] else self.grey))
        c.drawString(20*mm, y - 1*mm,
                     "Cargo insurance is taken out by Baker Hughes."
                     if sp["insurance"] else
                     "No cargo insurance obligation on Baker Hughes.")
        y -= 7*mm
        c.setFillColorRGB(*self.ink)
        y = self._wrap(c, sp["note"], 20*mm, y, self.W - 40*mm, leading=4.6*mm)
        return y

    def _incoterm_matrix_page(self, c, selected):
        self._running_head(c)
        self._title_block(c, "Incoterms 2020",
                          "the rule applied to this proposal is highlighted")
        n = len(INCOTERM_STAGES)
        x0 = 44*mm
        cw = (self.W - 20*mm - x0 - 18*mm) / n
        xi = x0 + n*cw + 9*mm
        y = self.H - 66*mm
        c.setFont("Helvetica", 6); c.setFillColorRGB(*self.grey)
        for i, (short, _l) in enumerate(INCOTERM_STAGES):
            c.drawCentredString(x0 + i*cw + cw/2, y, short)
        c.drawCentredString(xi, y, "Insurance")
        y -= 5*mm
        for code in INCOTERMS:
            sp = incoterm_spec(code)
            if not sp:
                continue
            on = (code == selected)
            if on:
                c.setFillColorRGB(*_rl_colour(BH_SEC))
                c.rect(20*mm, y - 1.6*mm, self.W - 40*mm, 5.6*mm, fill=1, stroke=0)
            c.setFont("Helvetica-Bold" if on else "Helvetica", 8)
            c.setFillColorRGB(*(self.green if on else self.ink))
            c.drawRightString(x0 - 3*mm, y, code)
            for i in range(n):
                x = x0 + i*cw + cw/2
                if sp["seller"][i]:
                    c.setFillColorRGB(*_rl_colour(BH_OK if on else BH_OK_PALE))
                    c.circle(x, y + 1*mm, 1.2*mm, fill=1, stroke=0)
                else:
                    c.setStrokeColorRGB(*_rl_colour(BH_GRID))
                    c.line(x - 1*mm, y + 1*mm, x + 1*mm, y + 1*mm)
            if sp["insurance"]:
                c.setFillColorRGB(*_rl_colour(BH_OK if on else BH_OK_PALE))
                c.circle(xi, y + 1*mm, 1.2*mm, fill=1, stroke=0)
            else:
                c.setStrokeColorRGB(*_rl_colour(BH_GRID))
                c.line(xi - 1*mm, y + 1*mm, xi + 1*mm, y + 1*mm)
            y -= 6.4*mm

        y -= 4*mm
        c.setFont("Helvetica", 7.5); c.setFillColorRGB(*self.grey)
        c.drawString(20*mm, y, "●  paid by Baker Hughes (seller)          "
                               "–  paid by the customer (buyer)")
        y -= 10*mm
        c.setFont("Helvetica-Bold", 9); c.setFillColorRGB(*self.ink)
        c.drawString(20*mm, y, "The eight steps of the chain")
        y -= 6*mm
        for short, long in INCOTERM_STAGES:
            c.setFont("Helvetica-Bold", 8); c.setFillColorRGB(*self.ink)
            c.drawString(20*mm, y, short)
            c.setFont("Helvetica", 8); c.setFillColorRGB(*self.grey)
            c.drawString(50*mm, y, long)
            y -= 5.2*mm
        self._running_foot(c)
        c.showPage()

    # ── plan ──────────────────────────────────────────────────────────────
    def _plan(self):
        """Ordered list of blocks to assemble."""
        plan = []
        for sec in self.struct:
            items = []
            for it, depth in _leaves(sec):
                if it["sub"]:
                    items.append(("group", it, depth, None))
                    continue
                st = self.status.get(it["code"], {})
                path = st.get("path", "")
                ref, title = it["ref"], st.get("title", it["title"])
                if it["code"] in ("0.1", "0.3") or st.get("state") == "tool":
                    kind = {"0.1": "cover", "0.3": "toc"}.get(it["code"], "tool")
                    items.append((kind, it, depth, None))
                elif path and path.lower().endswith(".pdf"):
                    items.append(("pdf", it, depth, path))
                elif path and path.lower().endswith(PDF_IMG_EXT):
                    items.append(("image", it, depth, path))
                elif path:
                    items.append(("attach", it, depth, path))
                else:
                    items.append(("missing", it, depth, None))
            if self.o["skip_empty"] and all(k == "missing" for k, *_ in items):
                continue
            plan.append((sec, items))
        return plan

    def _block_pages(self, kind, it, path):
        """How many pages a block will use (0 when it is skipped)."""
        if kind == "group":
            return 0
        if kind == "cover":
            return 1
        if kind == "toc":
            return 0                      # counted separately
        if kind == "tool":
            code = it["code"]
            if code == "0.2":
                return 1
            if code.startswith("2.1") and self.o["scope"]:
                return self._scope_page_count
            if code.startswith("11") and self.o["commercial"]:
                return 2
            return 1 if self.o["placeholders"] else 0
        if kind == "pdf":
            try:
                return len(PdfReader(path).pages)
            except Exception:
                return 1
        if kind in ("image", "attach"):
            return 1
        return 1 if self.o["placeholders"] else 0     # missing

    # ── build ─────────────────────────────────────────────────────────────
    def build(self, out_path):
        if not (HAS_REPORTLAB and HAS_PYPDF):
            raise RuntimeError("reportlab and pypdf are required.")
        plan = self._plan()

        # rough scope size, needed before the page map
        self._scope_page_count = max(1, (len(self.d.get("scope_of_supply") or []) + 38) // 39)

        # ── table of contents entries ──
        entries = []
        for sec, items in plan:
            if sec["num"] != 0:
                entries.append((0, f"Section {sec.get('disp', sec['num'])}", sec["title"]))
            for kind, it, depth, path in items:
                if sec["num"] == 0 and it["code"] in ("0.1", "0.3"):
                    continue
                if it["ref"] == str(sec.get("disp", sec["num"])):      # section = single document
                    continue
                entries.append((depth + (0 if sec["num"] == 0 else 1),
                                it["ref"], self.status.get(it["code"], {})
                                .get("title", it["title"])))
        toc_pages = max(1, (len(entries) + 33) // 34)

        # ── page map ──
        page, sec_page, item_page = 1, {}, {}
        for sec, items in plan:
            first = None
            for kind, it, depth, path in items:
                if kind == "cover":
                    item_page[it["code"]] = page; page += 1; continue
                if kind == "toc":
                    item_page[it["code"]] = page; page += toc_pages; continue
                if kind == "group":
                    item_page[it["code"]] = page; continue
                if first is None and sec["num"] != 0:
                    sec_page[sec["num"]] = page
                    page += 1                       # divider
                    first = page
                item_page[it["code"]] = page
                page += self._block_pages(kind, it, path)
            if first is None and sec["num"] != 0:
                sec_page[sec["num"]] = page; page += 1
        total = page - 1

        toc_nums = []
        for sec, items in plan:
            if sec["num"] != 0:
                toc_nums.append(sec_page.get(sec["num"]))
            for kind, it, depth, path in items:
                if sec["num"] == 0 and it["code"] in ("0.1", "0.3"):
                    continue
                if it["ref"] == str(sec.get("disp", sec["num"])):
                    continue
                toc_nums.append(item_page.get(it["code"]))

        # ── render ──
        writer = PdfWriter()
        tmp_dir = os.path.join(os.path.dirname(out_path) or ".", "._bog_tmp")
        os.makedirs(tmp_dir, exist_ok=True)
        seq = [0]

        def flush(cnv, buf_path):
            cnv.save()
            try:
                for pg in PdfReader(buf_path).pages:
                    writer.add_page(pg)
            except Exception:
                pass

        def new_canvas():
            seq[0] += 1
            bp = os.path.join(tmp_dir, f"gen_{seq[0]:03d}.pdf")
            return rl_canvas.Canvas(bp, pagesize=A4), bp

        for sec, items in plan:
            disp = sec.get("disp", sec["num"])
            self.section_starts.append(
                ("00_Front matter" if sec["num"] == 0
                 else f"{int(disp):02d}_{sec['title']}", len(writer.pages)))
            need_div = sec["num"] != 0 and any(
                k != "group" for k, *_ in items)
            if need_div:
                c, bp = new_canvas(); self._divider(c, sec); flush(c, bp)
            for kind, it, depth, path in items:
                ref = it["ref"]
                title = self.status.get(it["code"], {}).get("title", it["title"])
                if kind == "group":
                    continue
                if kind == "cover":
                    c, bp = new_canvas(); self._cover(c); flush(c, bp)
                    self.report.append((ref, title, "generated"))
                elif kind == "toc":
                    c, bp = new_canvas()
                    self._toc(c, entries, toc_nums); flush(c, bp)
                    self.report.append((ref, title, "generated"))
                elif kind == "tool":
                    c, bp = new_canvas()
                    code = it["code"]
                    if code == "0.2":
                        self._confidentiality(c)
                        self.report.append((ref, title, "generated"))
                    elif code.startswith("2.1") and self.o["scope"]:
                        n = self._scope_pages(c)
                        self._scope_page_count = n
                        self.report.append((ref, title, "generated from the form"))
                    elif code.startswith("11") and self.o["commercial"]:
                        self._commercial_page(c)
                        self.report.append((ref, title, "generated from the form"))
                    elif self.o["placeholders"]:
                        self._placeholder(c, ref, title,
                                          "This section is written directly in the "
                                          "proposal document.")
                        self.report.append((ref, title, "placeholder"))
                    flush(c, bp)
                elif kind == "pdf":
                    try:
                        for pg in PdfReader(path).pages:
                            writer.add_page(pg)
                        self.report.append((ref, title, os.path.basename(path)))
                    except Exception as e:
                        c, bp = new_canvas()
                        self._placeholder(c, ref, title,
                                          f"The file {os.path.basename(path)} could "
                                          f"not be read ({e}).")
                        flush(c, bp)
                        self.report.append((ref, title, "unreadable file"))
                elif kind == "image":
                    c, bp = new_canvas(); self._image_page(c, ref, title, path)
                    flush(c, bp)
                    self.report.append((ref, title, os.path.basename(path)))
                elif kind == "attach":
                    c, bp = new_canvas(); self._attachment_note(c, ref, title, path)
                    flush(c, bp)
                    self.report.append((ref, title,
                                        f"attached: {os.path.basename(path)}"))
                elif kind == "missing" and self.o["placeholders"]:
                    c, bp = new_canvas()
                    self._placeholder(c, ref, title,
                                      "This document is not yet available in the "
                                      "project folder nor in the standard database. "
                                      "Add it and generate the proposal again.")
                    flush(c, bp)
                    self.report.append((ref, title, "MISSING — placeholder"))
                elif kind == "missing":
                    self.report.append((ref, title, "MISSING — skipped"))

        # ── stamp headers and footers on every page ──
        if self.o["stamp"]:
            n = len(writer.pages)
            stamp_path = os.path.join(tmp_dir, "stamp.pdf")
            sc = rl_canvas.Canvas(stamp_path, pagesize=A4)
            for i in range(n):
                self._running_foot(sc, i + 1, n)
                sc.showPage()
            sc.save()
            try:
                stamps = PdfReader(stamp_path).pages
                for i, pg in enumerate(writer.pages):
                    if i == 0:
                        continue                      # cover keeps its own footer
                    try:
                        pg.merge_page(stamps[i])
                    except Exception:
                        pass
            except Exception:
                pass

        with open(out_path, "wb") as fh:
            writer.write(fh)
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return len(writer.pages)


# ═════════════════════════════════════════════════════════════════════════════
# DOCUMENT CONVERSION
# Anything attached to the proposal ends up as PDF pages. Office files go
# through Microsoft Office when it is installed (same rendering as a manual
# "Save as PDF"), LibreOffice otherwise; a workbook can also be drawn directly
# when neither is available.
# ═════════════════════════════════════════════════════════════════════════════
import base64
import html as _html
import re as _re
import tempfile
import zipfile

_NO_WINDOW = 0x08000000 if sys.platform.startswith("win") else 0
XL_EXT   = (".xlsx", ".xlsm", ".xls", ".xlsb")
WORD_EXT = (".docx", ".doc", ".docm", ".rtf")
PPT_EXT  = (".pptx", ".ppt")

_PS_EXCEL = r"""
$ErrorActionPreference = 'Stop'
$src = %s; $dst = %s
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try { $x.AutomationSecurity = 3 } catch {}
try { $x.AskToUpdateLinks = $false } catch {}
try {
  $wb = $x.Workbooks.Open($src, 0, $true)
  foreach ($ws in $wb.Worksheets) {
    if ($ws.Visible -eq -1) {
      try {
        $ps = $ws.PageSetup
        $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = 1
        try {
          if ($ps.PrintArea) { $ur = $ws.Range($ps.PrintArea) } else { $ur = $ws.UsedRange }
          if ($ur.Height -gt 2.4 * $ur.Width) { $ps.FitToPagesTall = $false }
          if ($ur.Width -gt $ur.Height) { $ps.Orientation = 2 } else { $ps.Orientation = 1 }
          $ps.CenterHorizontally = $true
        } catch {}
      } catch {}
    }
  }
  $wb.ExportAsFixedFormat(0, $dst)
  $wb.Close($false)
} finally {
  $x.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
"""

_PS_WORD = r"""
$ErrorActionPreference = 'Stop'
$src = %s; $dst = %s
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($src, $false, $true, $false)
  try { $d.ExportAsFixedFormat($dst, 17) } catch { $d.SaveAs2($dst, 17) }
  $d.Close($false)
} finally {
  $w.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
"""


def _ps_quote(path):
    return "'" + str(path).replace("'", "''") + "'"


def _run_powershell(script, timeout=300):
    exe = shutil.which("powershell") or shutil.which("pwsh")
    if not exe:
        return False, "PowerShell not available"
    try:
        # encoded: no quoting or line-break surprise on the command line
        blob = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        r = subprocess.run(
            [exe, "-NoProfile", "-NonInteractive", "-EncodedCommand", blob],
            capture_output=True, text=True, timeout=timeout,
            creationflags=_NO_WINDOW)
        return r.returncode == 0, (r.stderr or r.stdout or "").strip()[-300:]
    except Exception as e:
        return False, str(e)


_VBS_WORD = """
Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
Set d = w.Documents.Open(WScript.Arguments(0), False, True, False)
d.ExportAsFixedFormat WScript.Arguments(1), 17
d.Close False
w.Quit
"""

_VBS_EXCEL = """
Set x = CreateObject("Excel.Application")
x.Visible = False
x.DisplayAlerts = False
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
On Error Resume Next
For Each ws In wb.Worksheets
  If ws.Visible = -1 Then
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    Set ur = ws.UsedRange
    If Len(ws.PageSetup.PrintArea) > 0 Then Set ur = ws.Range(ws.PageSetup.PrintArea)
    If ur.Height > 2.4 * ur.Width Then ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.CenterHorizontally = True
    If ur.Width > ur.Height Then
      ws.PageSetup.Orientation = 2
    Else
      ws.PageSetup.Orientation = 1
    End If
  End If
Next
On Error GoTo 0
wb.ExportAsFixedFormat 0, WScript.Arguments(1)
wb.Close False
x.Quit
"""


def _run_vbs(script, src, dst, timeout=300):
    exe = shutil.which("cscript")
    if not exe:
        return False, "cscript not available"
    path = os.path.join(tempfile.gettempdir(), f"ybt_conv_{os.getpid()}.vbs")
    try:
        with open(path, "w", encoding="ascii") as fh:
            fh.write(script)
        r = subprocess.run([exe, "//nologo", "//B", path, src, dst],
                           capture_output=True, text=True, timeout=timeout,
                           creationflags=_NO_WINDOW)
        return r.returncode == 0, (r.stderr or r.stdout or "").strip()[-300:]
    except Exception as e:
        return False, str(e)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def _run_pywin32(src, dst, excel):
    """Office driven directly from Python, when pywin32 is installed."""
    try:
        import pythoncom
        import win32com.client as com
    except Exception:
        return False, "pywin32 not installed"
    app = None
    try:
        pythoncom.CoInitialize()
        if excel:
            app = com.DispatchEx("Excel.Application")
            app.Visible = False; app.DisplayAlerts = False
            wb = app.Workbooks.Open(src, 0, True)
            for ws in wb.Worksheets:
                if ws.Visible == -1:
                    try:
                        ps = ws.PageSetup
                        ur = ws.Range(ps.PrintArea) if ps.PrintArea else ws.UsedRange
                        ps.Zoom = False
                        ps.FitToPagesWide = 1
                        # one page per sheet, unless it would be unreadable
                        ps.FitToPagesTall = False if ur.Height > 2.4 * ur.Width else 1
                        ps.Orientation = 2 if ur.Width > ur.Height else 1
                        ps.CenterHorizontally = True
                    except Exception:
                        pass
            wb.ExportAsFixedFormat(0, dst)
            wb.Close(False)
        else:
            app = com.DispatchEx("Word.Application")
            app.Visible = False; app.DisplayAlerts = 0
            doc = app.Documents.Open(src, False, True, False)
            doc.ExportAsFixedFormat(dst, 17)
            doc.Close(False)
        return True, ""
    except Exception as e:
        return False, str(e)[-300:]
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


CONVERSION_LOG = []        # why each route failed, for the last file tried


def _find_soffice():
    forced = os.environ.get("YBTOOL_SOFFICE")
    if forced and os.path.exists(forced):
        return forced
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    for base in (os.environ.get("PROGRAMFILES", ""),
                 os.environ.get("PROGRAMFILES(X86)", "")):
        cand = os.path.join(base, "LibreOffice", "program", "soffice.exe")
        if base and os.path.exists(cand):
            return cand
    return None


def _soffice_to_pdf(src, dst):
    exe = _find_soffice()
    if not exe:
        return False
    out_dir = tempfile.mkdtemp(prefix="ybt_lo_")
    try:
        target = "pdf"
        if os.path.splitext(src)[1].lower() in XL_EXT:
            # one page per sheet rather than columns cut across pages
            target = ('pdf:calc_pdf_Export:{"SinglePageSheets":'
                      '{"type":"boolean","value":"true"}}')
        subprocess.run([exe, "--headless", "--convert-to", target,
                        "--outdir", out_dir, src],
                       capture_output=True, timeout=300,
                       creationflags=_NO_WINDOW)
        made = os.path.join(
            out_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if os.path.exists(made) and os.path.getsize(made) > 0:
            shutil.move(made, dst)
            if target != "pdf":
                _fit_pages_a4(dst)
            return True
    except Exception:
        pass
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)
    return False


def _fit_pages_a4(path):
    """Put free-size pages (one per sheet) onto A4, portrait or landscape,
    scaled to the margins and kept at the top of the page."""
    try:
        from pypdf import PageObject, Transformation
        reader, writer = PdfReader(path), PdfWriter()
        margin = 12 * mm
        for pg in reader.pages:
            w, h = float(pg.mediabox.width), float(pg.mediabox.height)
            x0, y0 = float(pg.mediabox.left), float(pg.mediabox.bottom)
            if w < 4 or h < 4 or _page_blank(pg):
                continue                    # empty sheet: no page
            # landscape only when the sheet does not fit across a portrait page
            wide = w > h * 1.15 and w > A4[0] - 2 * margin
            pw, ph = (A4[1], A4[0]) if wide else A4
            k = min((pw - 2 * margin) / w, (ph - 2 * margin) / h, 1.5)
            new = PageObject.create_blank_page(width=pw, height=ph)
            op = (Transformation().translate(-x0, -y0).scale(k, k)
                  .translate((pw - w * k) / 2, ph - margin - h * k))
            new.merge_transformed_page(pg, op)
            writer.add_page(new)
        if not len(writer.pages):
            return
        tmp = path + ".fit"
        with open(tmp, "wb") as fh:
            writer.write(fh)
        os.replace(tmp, path)
    except Exception:
        pass


def _office_to_pdf(src, dst):
    """Convert a Word or Excel file. Returns the engine used, or None.
    Microsoft Office is tried three ways before LibreOffice, because what
    is allowed differs from one company computer to the next."""
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    ext = os.path.splitext(src)[1].lower()
    del CONVERSION_LOG[:]

    def made():
        if not (os.path.exists(dst) and os.path.getsize(dst) > 0):
            return False
        try:
            if all(_page_blank(p) for p in PdfReader(dst).pages):
                CONVERSION_LOG.append("PDF produced, but every page is blank")
                return False
        except Exception:
            pass
        return True

    if sys.platform.startswith("win") and ext in XL_EXT + WORD_EXT:
        excel = ext in XL_EXT
        name = "Microsoft Excel" if excel else "Microsoft Word"
        routes = [
            ("Office via Python", lambda: _run_pywin32(src, dst, excel)),
            ("Office via PowerShell", lambda: _run_powershell(
                (_PS_EXCEL if excel else _PS_WORD)
                % (_ps_quote(src), _ps_quote(dst)))),
            ("Office via script host", lambda: _run_vbs(
                _VBS_EXCEL if excel else _VBS_WORD, src, dst)),
        ]
        for label, run in routes:
            ok, msg = run()
            if ok and made():
                return name
            CONVERSION_LOG.append(f"{label}: {msg or 'no PDF produced'}")
    if _soffice_to_pdf(src, dst) and made():
        return "LibreOffice"
    CONVERSION_LOG.append("LibreOffice: not installed" if not _find_soffice()
                          else "LibreOffice: no PDF produced")
    return None


def _xlsx_to_pdf_basic(src, dst):
    """Draw a workbook without Office: cell values, fills, bold, merged cells
    and borders, each sheet scaled to the page width."""
    wb = openpyxl.load_workbook(src, data_only=True)
    c = None
    for ws in wb.worksheets:
        if ws.sheet_state != "visible" or ws.max_row < 1:
            continue
        r1, r2, c1, c2 = None, 0, None, 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.value not in (None, ""):
                    r1 = cell.row if r1 is None else min(r1, cell.row)
                    r2 = max(r2, cell.row)
                    c1 = cell.column if c1 is None else min(c1, cell.column)
                    c2 = max(c2, cell.column)
        if r1 is None:
            continue
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= r2 and mr.min_col <= c2:
                r2, c2 = max(r2, mr.max_row), max(c2, mr.max_col)
        letter = openpyxl.utils.get_column_letter
        cw = [(ws.column_dimensions[letter(j)].width or 8.43) * 5.6
              for j in range(c1, c2 + 1)]
        rh = [(ws.row_dimensions[i].height or 15.0) for i in range(r1, r2 + 1)]
        land = sum(cw) > 520
        pw, ph = (A4[1], A4[0]) if land else A4
        margin = 12 * mm
        k = min(1.0, (pw - 2 * margin) / max(sum(cw), 1),
                (ph - 2 * margin) / max(sum(rh), 1))       # one sheet, one page
        cw, rh = [w * k for w in cw], [h * k for h in rh]
        xs = [margin]
        for w in cw:
            xs.append(xs[-1] + w)
        merged, hidden = {}, set()
        for mr in ws.merged_cells.ranges:
            merged[(mr.min_row, mr.min_col)] = (mr.max_row, mr.max_col)
            for i in range(mr.min_row, mr.max_row + 1):
                for j in range(mr.min_col, mr.max_col + 1):
                    if (i, j) != (mr.min_row, mr.min_col):
                        hidden.add((i, j))
        if c is None:
            c = rl_canvas.Canvas(dst, pagesize=(pw, ph))
        else:
            c.setPageSize((pw, ph))
        y = ph - margin
        for ri, i in enumerate(range(r1, r2 + 1)):
            if y - rh[ri] < margin:
                c.showPage(); c.setPageSize((pw, ph)); y = ph - margin
            for ci, j in enumerate(range(c1, c2 + 1)):
                if (i, j) in hidden:
                    continue
                cell = ws.cell(row=i, column=j)
                x0, x1, h = xs[ci], xs[ci + 1], rh[ri]
                if (i, j) in merged:
                    mr_, mc_ = merged[(i, j)]
                    x1 = xs[min(mc_, c2) - c1 + 1]
                    h = sum(rh[ri:ri + (min(mr_, r2) - i) + 1])
                fill = getattr(cell.fill, "fgColor", None)
                rgb = getattr(fill, "rgb", None)
                if (cell.fill is not None and cell.fill.fill_type == "solid"
                        and isinstance(rgb, str) and len(rgb) == 8
                        and rgb[2:].upper() != "FFFFFF"):
                    c.setFillColorRGB(*_rl_colour(rgb[2:]))
                    c.rect(x0, y - h, x1 - x0, h, fill=1, stroke=0)
                b = cell.border
                if b is not None and any(getattr(s, "style", None) for s in
                                         (b.left, b.right, b.top, b.bottom)):
                    c.setStrokeColorRGB(0.55, 0.57, 0.57); c.setLineWidth(0.4)
                    c.rect(x0, y - h, x1 - x0, h, fill=0, stroke=1)
                v = cell.value
                if v in (None, ""):
                    continue
                if isinstance(v, float):
                    v = f"{v:,.4f}".rstrip("0").rstrip(".")
                txt = str(v)
                size = max(4.5, min(11, (cell.font.sz or 10) * k))
                face = "Helvetica-Bold" if cell.font and cell.font.b else "Helvetica"
                fc = getattr(getattr(cell.font, "color", None), "rgb", None)
                if isinstance(fc, str) and len(fc) == 8:
                    c.setFillColorRGB(*_rl_colour(fc[2:]))
                else:
                    c.setFillColorRGB(0.1, 0.12, 0.12)
                c.setFont(face, size)
                room = x1 - x0 - 3
                while txt and c.stringWidth(txt, face, size) > room and len(txt) > 1:
                    txt = txt[:-1]
                ty = y - h / 2 - size * 0.35
                align = (cell.alignment.horizontal or
                         ("right" if isinstance(cell.value, (int, float)) else "left"))
                if align in ("center", "centerContinuous"):
                    c.drawCentredString((x0 + x1) / 2, ty, txt)
                elif align == "right":
                    c.drawRightString(x1 - 1.5, ty, txt)
                else:
                    c.drawString(x0 + 1.5, ty, txt)
            y -= rh[ri]
        c.showPage()
    if c is None:
        raise ValueError("empty workbook")
    c.save()


def _image_to_pdf(src, dst):
    img = ImageReader(src)
    iw, ih = img.getSize()
    pw, ph = (A4[1], A4[0]) if iw > ih * 1.15 else A4
    k = min((pw - 20 * mm) / iw, (ph - 20 * mm) / ih)
    c = rl_canvas.Canvas(dst, pagesize=(pw, ph))
    c.drawImage(img, (pw - iw * k) / 2, (ph - ih * k) / 2, iw * k, ih * k,
                preserveAspectRatio=True, mask="auto")
    c.showPage(); c.save()


# ── Excel made printable before conversion ──────────────────────────────────
# A sheet whose formatting runs far beyond its data (borders down to row
# 1000, a coloured column…) is shrunk by "fit to one page" until it looks
# empty. The range that really holds something becomes the print area, and
# sheets that hold nothing are hidden. The workbook itself is not touched.
import zipfile as _zipfile, html as _html_x

def _col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_txt(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _xlsx_extent(files, sheet_path, shared):
    """(r1, c1, r2, c2) of what is really on a sheet: cells holding a value,
    merged blocks, pictures and charts. Cells that only carry formatting are
    ignored: they are what makes Excel shrink a sheet to an empty-looking
    page. None when the sheet shows nothing."""
    xml = files.get(sheet_path, b"").decode("utf-8", "ignore")
    rows, cols = [], []

    def add(r, c):
        rows.append(r); cols.append(c)

    for m in re.finditer(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S):
        attrs, body = m.group(1), m.group(2) or ""
        ref = re.search(r'\br="([A-Z]+)(\d+)"', attrs)
        if not ref or not body:
            continue
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        t = re.search(r'\bt="(\w+)"', attrs)
        if "<is>" in body:
            txt = re.sub(r"<[^>]+>", "", body)
        elif v is None:
            continue
        elif t and t.group(1) == "s":
            try:
                txt = shared[int(v.group(1))]
            except (ValueError, IndexError):
                txt = "x"
        else:
            txt = v.group(1)
        if txt.strip():
            add(int(ref.group(2)), _col_num(ref.group(1)))
    for m in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', xml):
        r1, c1 = int(m.group(2)), _col_num(m.group(1))
        if r1 in rows and any(c == c1 for r, c in zip(rows, cols) if r == r1):
            add(int(m.group(4)), _col_num(m.group(3)))
    # pictures and charts anchored on the sheet
    rel_path = os.path.dirname(sheet_path) + "/_rels/" + os.path.basename(sheet_path) + ".rels"
    rels = files.get(rel_path, b"").decode("utf-8", "ignore")
    for tgt in re.findall(r'Target="([^"]*drawings/[^"]+\.xml)"', rels):
        dpath = os.path.normpath(os.path.join(os.path.dirname(sheet_path), tgt)).replace("\\", "/")
        dxml = files.get(dpath, b"").decode("utf-8", "ignore")
        for a in re.finditer(r"<xdr:(from|to)>\s*<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>",
                             dxml, re.S):
            add(int(a.group(3)) + 1, int(a.group(2)) + 1)
    if not rows:
        return None
    return min(rows), min(cols), max(rows), max(cols)


def xlsx_prepare(src, dst):
    """Copy of an .xlsx/.xlsm ready to print: each sheet prints only the
    range that holds something (unless the author set a print area), and
    sheets that hold nothing are hidden so they do not come out as blank
    pages. Returns (path_to_convert, notes). Every other part of the file
    is copied byte for byte."""
    with _zipfile.ZipFile(src) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i.filename) for i in infos}
    wb = files["xl/workbook.xml"].decode("utf-8")
    rels = files.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
    rid_target = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
    rid_target.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    shared = []
    if "xl/sharedStrings.xml" in files:
        ss = files["xl/sharedStrings.xml"].decode("utf-8", "ignore")
        for si in re.findall(r"<si>(.*?)</si>", ss, re.S):
            shared.append(_html_x.unescape(re.sub(r"<[^>]+>", "", si)))
    sheets = re.findall(r"<sheet\b[^>]*/>", wb)
    has_area = set()
    for m in re.finditer(r'<definedName\b([^>]*)>(.*?)</definedName>', wb, re.S):
        if 'name="_xlnm.Print_Area"' in m.group(1):
            lid = re.search(r'localSheetId="(\d+)"', m.group(1))
            if lid:
                has_area.add(int(lid.group(1)))
    notes, areas, shown = [], [], []
    new_sheets = []
    for idx, tag in enumerate(sheets):
        name = _html_x.unescape(re.search(r'name="([^"]*)"', tag).group(1))
        rid = re.search(r'r:id="([^"]+)"', tag) or re.search(r'\bid="([^"]+)"', tag)
        tgt = rid_target.get(rid.group(1), "") if rid else ""
        path = tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
        hidden = re.search(r'state="(hidden|veryHidden)"', tag)
        ext = _xlsx_extent(files, path, shared) if path in files else (1, 1, 1, 1)
        if hidden:
            new_sheets.append(tag); continue
        if ext is None:
            notes.append(f"sheet '{name}' is empty: left out")
            new_sheets.append(re.sub(r'\s+state="\w+"', "", tag)
                              .replace("/>", ' state="hidden"/>'))
            continue
        shown.append(idx)
        new_sheets.append(tag)
        if idx not in has_area:
            r1, c1, r2, c2 = ext
            q = "'" + name.replace("'", "''") + "'"
            areas.append((idx, f"{q}!${_col_txt(c1)}${r1}:${_col_txt(c2)}${r2}"))
    if not shown:
        raise ValueError("the workbook holds no value on any visible sheet")
    for old, new in zip(sheets, new_sheets):
        if old != new:
            wb = wb.replace(old, new, 1)
    if areas:
        names = "".join(f'<definedName name="_xlnm.Print_Area" localSheetId="{i}">'
                        f"{_html_x.escape(a, quote=False)}</definedName>" for i, a in areas)
        if "<definedNames>" in wb:
            wb = wb.replace("<definedNames>", "<definedNames>" + names, 1)
        elif "<definedNames/>" in wb:
            wb = wb.replace("<definedNames/>", f"<definedNames>{names}</definedNames>", 1)
        else:
            anchor = re.search(r"</sheets>", wb)
            wb = wb[:anchor.end()] + f"<definedNames>{names}</definedNames>" + wb[anchor.end():]
    # the sheet that opens first must be one that is shown
    wb = re.sub(r'activeTab="\d+"', f'activeTab="{shown[0]}"', wb)
    wb = re.sub(r'firstSheet="\d+"', f'firstSheet="{shown[0]}"', wb)
    if not areas and not notes:
        return src, notes
    files["xl/workbook.xml"] = wb.encode("utf-8")
    with _zipfile.ZipFile(dst, "w", _zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, files[i.filename])
    return dst, notes


# ── blank pages ──────────────────────────────────────────────────────────────
_PAINT_RE = re.compile(rb"(?<![A-Za-z*'\"])(Tj|TJ|Do|sh|BI|S|s|B\*?|b\*?|f\*?|F)(?![A-Za-z*])")
_COLOUR_RE = re.compile(rb"((?:-?[\d.]+\s+){1,4})(rg|g|k|sc|scn)(?![A-Za-z])")


def _page_blank(pg):
    """True when a page would print nothing: no size, no text, no picture,
    no line, only white fills. Read from the page itself, no renderer."""
    try:
        w, h = float(pg.mediabox.width), float(pg.mediabox.height)
        if w < 4 or h < 4:
            return True
        if (pg.extract_text() or "").strip():
            return False
        cont = pg.get_contents()
        try:
            data = cont.get_data() if cont else b""
        except Exception:
            data = b""
        if not data.strip():
            return True
        ops = set(m.group(1) for m in _PAINT_RE.finditer(data))
        if ops & {b"Tj", b"TJ", b"Do", b"sh", b"BI", b"S", b"s",
                  b"B", b"B*", b"b", b"b*"}:
            return False
        if not ops:
            return True
        # only fills: blank if every colour set is white
        cols = _COLOUR_RE.findall(data)
        if not cols:
            return False                    # default fill colour is black
        for nums, op in cols:
            v = [float(x) for x in nums.split()]
            white = (all(x == 0 for x in v) if op == b"k"
                     else all(x >= 0.99 for x in v))
            if not white:
                return False
        return True
    except Exception:
        return False


def pdf_drop_blank(src, dst=None):
    """Copy of src without its blank pages. Returns (path, kept, dropped);
    path is src itself when nothing had to go."""
    try:
        reader = PdfReader(src)
        keep = [p for p in reader.pages if not _page_blank(p)]
        dropped = len(reader.pages) - len(keep)
        if not dropped or not keep:
            return src, len(keep), dropped
        w = PdfWriter()
        for p in keep:
            w.add_page(p)
        dst = dst or src
        tmp = dst + ".nb"
        with open(tmp, "wb") as fh:
            w.write(fh)
        os.replace(tmp, dst)
        return dst, len(keep), dropped
    except Exception:
        return src, 1, 0


def to_pdf(path, tmp_dir, cache=None):
    """Return (pdf_path, how). pdf_path is None when the file cannot be
    turned into pages."""
    if not path or not os.path.exists(path):
        return None, "file not found"
    ext = os.path.splitext(path)[1].lower()
    key = (os.path.abspath(path), os.path.getmtime(path))
    if ext == ".pdf":
        if cache is not None and key in cache:
            return cache[key]
        os.makedirs(tmp_dir, exist_ok=True)
        clean = os.path.join(tmp_dir, f"nb_{abs(hash(key)) % 10**10}_"
                             + _re.sub(r"[^A-Za-z0-9.]+", "_", os.path.basename(path)))
        out, kept, dropped = pdf_drop_blank(path, clean)
        if not kept:
            res = (None, "the PDF has no printed content")
        else:
            res = (out, f"{dropped} blank page(s) left out" if dropped else "")
        if cache is not None:
            cache[key] = res
        return res
    if cache is not None and key in cache:
        return cache[key]
    os.makedirs(tmp_dir, exist_ok=True)
    dst = os.path.join(
        tmp_dir, f"conv_{abs(hash(key)) % 10**10}_"
                 f"{_re.sub(r'[^A-Za-z0-9]+', '_', os.path.basename(path))}.pdf")
    res = (None, "format not supported")
    try:
        if ext in PDF_IMG_EXT and HAS_REPORTLAB:
            _image_to_pdf(path, dst); res = (dst, "image")
        elif ext in XL_EXT + WORD_EXT + PPT_EXT:
            src, notes = path, []
            if ext in (".xlsx", ".xlsm"):
                try:
                    src, notes = xlsx_prepare(
                        path, os.path.join(tmp_dir, "prn_" + os.path.basename(dst)
                                           .replace(".pdf", ext)))
                except ValueError as e:
                    raise RuntimeError(str(e))
                except Exception:
                    src, notes = path, []       # unusual file: as it is
            engine = _office_to_pdf(src, dst)
            if engine:
                _p, kept, dropped = pdf_drop_blank(dst)
                if dropped:
                    notes.append(f"{dropped} blank page(s) left out")
                res = (dst, f"converted with {engine}"
                       + (f" — {'; '.join(notes)}" if notes else ""))
            elif ext in (".xlsx", ".xlsm") and HAS_REPORTLAB:
                _xlsx_to_pdf_basic(path, dst)
                res = (dst, "drawn by the tool (no Office found)")
            else:
                res = (None, "needs Microsoft Office or LibreOffice")
    except Exception as e:
        res = (None, f"conversion failed ({e})")
    if cache is not None:
        cache[key] = res
    return res


def merge_pdfs(paths, dst):
    w = PdfWriter()
    for p in paths:
        for pg in PdfReader(p).pages:
            w.add_page(pg)
    with open(dst, "wb") as fh:
        w.write(fh)
    return dst


def _safe_name(text):
    return _re.sub(r"\s+", " ", _re.sub(r'[\\/:*?"<>|]+', " ", str(text))).strip()


def split_pdf_by_section(pdf_path, starts, folder):
    """starts: [(label, first_page_index)], ascending. One file per entry."""
    reader = PdfReader(pdf_path)
    n = len(reader.pages)
    os.makedirs(folder, exist_ok=True)
    made = []
    for i, (label, first) in enumerate(starts):
        last = starts[i + 1][1] if i + 1 < len(starts) else n
        if last <= first:
            continue
        w = PdfWriter()
        for k in range(first, last):
            w.add_page(reader.pages[k])
        out = os.path.join(folder, _safe_name(label) + ".pdf")
        with open(out, "wb") as fh:
            w.write(fh)
        made.append(out)
    return made


# ═════════════════════════════════════════════════════════════════════════════
# WORD SKELETON
# Technical_Proposal.docx, kept next to the tool, is the master layout. The
# builder works on a copy: it fills the project fields, puts the sections in
# the order chosen in the document manager, renumbers the headings, rewrites
# the table of contents, then converts to PDF and drops every attached
# document on its own pages right after the heading it belongs to.
# ═════════════════════════════════════════════════════════════════════════════
SKELETON_NAME = "Technical_Proposal.docx"

_XT  = _re.compile(r'<(/?)([A-Za-z0-9]+:[A-Za-z0-9]+)((?:\s[^<>]*?)?)(/?)>')
_WT  = _re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>')
_WTT = _re.compile(r'<w:t(?:\s[^>]*)?>([^<]*)</w:t>|<w:tab/>')
_BRK = _re.compile(r'<w:br w:type="page"/>')
_SHAPES = ("mc:AlternateContent", "w:drawing", "w:pict", "w:object")
_P_OPEN = _re.compile(r'<w:p(?:\s[^>]*)?>')


def _dx_top(xml):
    """Top-level elements of an XML fragment, as strings."""
    out, depth, start = [], 0, 0
    for m in _XT.finditer(xml):
        close, _name, _a, selfc = m.groups()
        if selfc:
            if depth == 0:
                out.append(xml[m.start():m.end()])
        elif close:
            depth -= 1
            if depth == 0:
                out.append(xml[start:m.end()])
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return out


def _dx_spans(xml, names):
    spans, depth, start = [], 0, 0
    for m in _XT.finditer(xml):
        close, name, _a, selfc = m.groups()
        if name not in names:
            continue
        if selfc:
            if depth == 0:
                spans.append((m.start(), m.end()))
        elif close:
            depth -= 1
            if depth == 0:
                spans.append((start, m.end()))
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return spans


def _dx_is_p(x):
    return x.startswith("<w:p>") or x.startswith("<w:p ")


def _dx_open(x):
    if x.endswith("/>") and "</w:p>" not in x:
        return x[:-2] + "></w:p>"
    return x


def _dx_text(x):
    """Visible text of a paragraph, shapes excluded, tabs read as spaces."""
    spans = _dx_spans(x, _SHAPES)
    out = []
    for m in _WTT.finditer(x):
        if any(a <= m.start() < b for a, b in spans):
            continue
        out.append(" " if m.group(1) is None else _html.unescape(m.group(1)))
    return "".join(out)


def _dx_all_text(x):
    return _html.unescape("".join(_WT.findall(x)))


def _dx_sub(x, a, b, new, drop_tabs=False):
    """Replace characters [a, b) of the paragraph text held in its own
    <w:t> nodes (shapes are left alone). The first node touched keeps its
    formatting and receives the new text."""
    if not drop_tabs and _dx_text_raw(x)[a:b] == new:
        return x                            # nothing to change
    spans = _dx_spans(x, _SHAPES)
    state = dict(pos=0, done=False)
    targets = [m for m in _WT.finditer(x)
               if not any(s <= m.start() < e for s, e in spans)]
    last = targets[-1].start() if targets else -1

    def repl(m):
        if any(s <= m.start() < e for s, e in spans):
            return m.group(0)
        t = _html.unescape(m.group(1))
        n0, n1 = state["pos"], state["pos"] + len(t)
        state["pos"] = n1
        hit = (n0 < b and n1 > a) or (a == b and n0 <= a <= n1)
        if not hit and not (m.start() == last and not state["done"]):
            return m.group(0)
        left = t[:max(0, min(len(t), a - n0))]
        right = t[max(0, min(len(t), b - n0)):]
        if not state["done"]:
            state["done"] = True
            t2 = left + new + right if hit else t + new
        else:
            t2 = left + right
        return ('<w:t xml:space="preserve">'
                + _html.escape(t2, quote=False) + '</w:t>')

    x = _WT.sub(repl, x)
    if drop_tabs:
        spans = _dx_spans(x, _SHAPES)
        parts, cur = [], 0
        for m in _re.finditer(r"<w:tab/>", x):
            if any(s <= m.start() < e for s, e in spans):
                continue
            parts.append(x[cur:m.start()]); cur = m.end()
        x = "".join(parts) + x[cur:]
    return x


def _dx_set_text(x, text):
    return _dx_sub(x, 0, len(_dx_text_raw(x)), text, drop_tabs=True)


def _dx_text_raw(x):
    spans = _dx_spans(x, _SHAPES)
    return "".join(_html.unescape(m.group(1)) for m in _WT.finditer(x)
                   if not any(a <= m.start() < b for a, b in spans))


def _dx_ppr(x):
    m = _P_OPEN.match(x)
    if not m:
        return None
    if x.startswith("<w:pPr/>", m.end()):
        return m.end(), m.end() + 8
    if x.startswith("<w:pPr>", m.end()):
        for a, b in _dx_spans(x, ("w:pPr",)):
            if a == m.end():
                return a, b
    return None


def _dx_style(x):
    sp = _dx_ppr(x)
    if not sp:
        return ""
    m = _re.search(r'<w:pStyle w:val="([^"]+)"', x[sp[0]:sp[1]])
    return m.group(1) if m else ""


def _dx_has_pbb(x):
    sp = _dx_ppr(x)
    return bool(sp) and "<w:pageBreakBefore/>" in x[sp[0]:sp[1]]


def _dx_set_pbb(x):
    x = _dx_open(x)
    if _dx_has_pbb(x):
        return x
    sp = _dx_ppr(x)
    if sp is None:
        m = _P_OPEN.match(x)
        return (x[:m.end()] + "<w:pPr><w:pageBreakBefore/></w:pPr>"
                + x[m.end():])
    a, b = sp
    ppr = x[a:b]
    if ppr == "<w:pPr/>":
        new = "<w:pPr><w:pageBreakBefore/></w:pPr>"
    else:
        m = _re.match(r"<w:pPr>(?:<w:pStyle[^>]*/>)?(?:<w:keepNext[^>]*/>)?"
                      r"(?:<w:keepLines[^>]*/>)?", ppr)
        new = ppr[:m.end()] + "<w:pageBreakBefore/>" + ppr[m.end():]
    return x[:a] + new + x[b:]


def _dx_drop_pbb(x):
    sp = _dx_ppr(x)
    if not sp:
        return x
    return x[:sp[0]] + x[sp[0]:sp[1]].replace("<w:pageBreakBefore/>", "") + x[sp[1]:]


def _dx_drop_numpr(x):
    sp = _dx_ppr(x)
    if not sp:
        return x
    ppr = _re.sub(r"<w:numPr>.*?</w:numPr>", "", x[sp[0]:sp[1]], flags=_re.S)
    return x[:sp[0]] + ppr + x[sp[1]:]


def _dx_has_numpr(x):
    sp = _dx_ppr(x)
    return bool(sp) and "<w:numPr>" in x[sp[0]:sp[1]]


def _dx_fresh(x):
    """Clone-safe copy: Word wants paragraph ids to be unique."""
    x = _re.sub(r"<w:bookmark(?:Start|End)\b[^>]*/>", "", x)
    return _re.sub(r'\sw14:(?:paraId|textId)="[^"]*"', "", x)


def _dx_visible(x):
    if any(tag in x for tag in ("<w:drawing", "<w:pict", "<w:object",
                                "<mc:AlternateContent", "<w:fldChar",
                                "<w:sectPr")):
        return True
    return bool(_dx_all_text(x).strip())


def _dx_filler(x):
    return _dx_is_p(x) and not _dx_visible(x)


def _dx_split_breaks(x, keep_mark=False):
    """One paragraph in, one or more out: every explicit page break becomes a
    'page break before' on what follows it. Same rendering, but blocks can
    then be moved without dragging a break along.

    Returns (paragraphs, carry). carry is True when the break closed the
    paragraph, so the next element has to open the new page. keep_mark
    follows the older Word layout, where the paragraph mark after a break
    drops onto the new page as an empty line."""
    spans = _dx_spans(x, _SHAPES)
    if not any(not any(a <= m.start() < b for a, b in spans)
               for m in _BRK.finditer(x)):
        return [x], False
    x = _dx_open(x)
    m = _P_OPEN.match(x)
    open_tag = m.group(0)
    sp = _dx_ppr(x)
    ppr = x[sp[0]:sp[1]] if sp else ""
    inner = x[(sp[1] if sp else m.end()):-len("</w:p>")]
    groups = [[]]
    for kid in _dx_top(inner):
        if kid.startswith("<w:r>") or kid.startswith("<w:r "):
            ks = _dx_spans(kid, _SHAPES)
            brs = [b for b in _BRK.finditer(kid)
                   if not any(a <= b.start() < e for a, e in ks)]
            if brs:
                rm = _re.match(r"<w:r(?:\s[^>]*)?>(?:<w:rPr>.*?</w:rPr>)?",
                               kid, _re.S)
                head, body = rm.group(0), kid[rm.end():-len("</w:r>")]
                cur = 0
                for i, b in enumerate(brs):
                    part = body[cur:b.start() - rm.end()]
                    if part.strip():
                        groups[-1].append(head + part + "</w:r>")
                    groups.append([])
                    cur = b.end() - rm.end()
                tail = body[cur:]
                if tail.strip():
                    groups[-1].append(head + tail + "</w:r>")
                continue
        groups[-1].append(kid)
    # a bookmark stays whole, on the part that carries the text
    marks = [k for g in groups for k in g if k.startswith("<w:bookmark")]
    if marks and len(groups) > 1:
        groups = [[k for k in g if not k.startswith("<w:bookmark")] for g in groups]
        texted = [i for i, g in enumerate(groups) if _WT.search("".join(g))]
        home = groups[texted[-1] if texted else -1]
        home[:0] = [k for k in marks if k.startswith("<w:bookmarkStart")]
        home.extend(k for k in marks if k.startswith("<w:bookmarkEnd"))
    out, carry = [], False
    for i, g in enumerate(groups):
        content = "".join(g)
        shown = _dx_visible("<w:p>" + content + "</w:p>")
        last = i == len(groups) - 1
        if not shown and (i == 0 or (last and not keep_mark)):
            if i > 0:
                carry = True                # break was the last thing here
            elif len(groups) == 1:
                out.append(x)
            continue
        tag = open_tag if not out else _dx_fresh(open_tag)
        p = tag + ppr + content + "</w:p>"
        if not shown:
            p = _dx_drop_numpr(p)
        if i > 0:
            p = _dx_set_pbb(p)
        out.append(p)
    if not out:                             # a paragraph holding only a break
        out, carry = [_dx_drop_numpr(open_tag + ppr + "</w:p>")], True
        if keep_mark:
            out, carry = [_dx_set_pbb(out[0])], False
    return out, carry


def _title_tokens(text):
    words = _re.findall(r"[a-z0-9]+", str(text).lower().replace("&", " "))
    stop = {"and", "of", "the", "for", "a", "an", "icl", "section"}
    out = set()
    for w in words:
        if w in stop or w.isdigit():
            continue
        out.add(w[:-1] if len(w) > 4 and w.endswith("s") else w)
    return out


def _title_sim(a, b):
    ta, tb = _title_tokens(a), _title_tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


class _AskBuiltin(Exception):
    """The Word master could not be filled: the user decides what happens."""
    def __init__(self, err, filled):
        super().__init__(str(err))
        self.err, self.filled = err, filled


class WordProposal:
    """Builds the proposal from the Word skeleton."""

    SLOT, SECM = "YBSLOT", "YBSEC"
    OFFSET_LINES = 10              # blank lines above a divider title

    def __init__(self, skeleton, data, struct, status, options=None,
                 logo=None, tmp_dir=None, title_of=None):
        self.skeleton = skeleton
        self.d = data
        self.struct = struct
        self.status = status
        self.o = dict(placeholders=True, scope=True, commercial=True,
                      skip_empty=False, frame=True)
        self.o.update(options or {})
        self.logo = logo
        self.tmp = tmp_dir or tempfile.mkdtemp(prefix="ybt_")
        self.title_of = title_of or (lambda it: it["title"])
        self.report = []
        self.section_starts = []
        self.engine = ""
        self._slots = {}           # marker -> [pdf paths]
        self._secs = {}            # marker -> label
        self._legacy = None

    # ── skeleton parsing ─────────────────────────────────────────────────
    def _load(self):
        with zipfile.ZipFile(self.skeleton) as z:
            self.names = [i.filename for i in z.infolist()]
            self.files = {n: z.read(n) for n in self.names}
        doc = self.files["word/document.xml"].decode("utf-8")
        a = doc.index("<w:body>") + len("<w:body>")
        b = doc.rindex("</w:body>")
        self.pre, body, self.post = doc[:a], doc[a:b], doc[b:]
        body = _re.sub(r"<w:lastRenderedPageBreak/>", "", body)
        body = _re.sub(r"<w:proofErr\b[^>]*/>", "", body)
        els = _dx_top(body)
        # a bookmark that opens in one block and closes in another cannot
        # follow a block that moves; those inside a single paragraph stay,
        # which keeps every table-of-contents link alive
        where = {}
        for i, x in enumerate(els):
            for m in _re.finditer(r'<w:bookmark(Start|End)\b[^>]*?w:id="(\d+)"', x):
                where.setdefault(m.group(2), {})[m.group(1)] = (
                    i, x.startswith("<w:bookmark"))
        loose = {k for k, v in where.items()
                 if len(v) < 2 or v["Start"][0] != v["End"][0]
                 or v["Start"][1] or v["End"][1]}
        if loose:
            rx = _re.compile(r'<w:bookmark(?:Start|End)\b[^>]*?w:id="(\d+)"[^>]*/>')
            els = [y for y in (rx.sub(lambda m: "" if m.group(1) in loose
                                      else m.group(0), x) for x in els) if y]
        self.tail = []
        if els and els[-1].startswith("<w:sectPr"):
            self.tail = [els.pop()]
        geo = "".join(self.tail)
        def twips(tag, attr, default):
            m = _re.search(rf'<w:{tag}\b[^>]*?w:{attr}="(\d+)"', geo)
            return (int(m.group(1)) if m else default) / 20.0
        self.page = (twips("pgSz", "w", 11906), twips("pgSz", "h", 16838))
        self.margins = dict(top=twips("pgMar", "top", 1440),
                            bottom=twips("pgMar", "bottom", 1440),
                            left=twips("pgMar", "left", 1440),
                            right=twips("pgMar", "right", 1440))
        settings = self.files.get("word/settings.xml", b"").decode("utf-8", "ignore")
        self.keep_mark = "<w:splitPgBreakAndParaMark" in settings
        styles = self.files.get("word/styles.xml", b"").decode("utf-8", "ignore")
        self.h1 = {"Titre1", "Heading1"}
        self.h2 = {"Titre2", "Heading2"}
        for m in _re.finditer(
                r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="([^"]+)"/>'
                r'(?:<w:basedOn w:val="([^"]+)"/>)?', styles):
            sid, name, base = m.groups()
            if name.lower() == "heading 2":
                self.h2.add(sid)
            if name.lower() == "heading 1":
                self.h1.add(sid)
        for m in _re.finditer(
                r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="[^"]+"/>'
                r'<w:basedOn w:val="([^"]+)"/>', styles):
            if m.group(2) in self.h2:
                self.h2.add(m.group(1))
        m = _re.search(r'<w:style [^>]*w:styleId="Titre1".*?</w:style>', styles, _re.S) \
            or _re.search(r'<w:style [^>]*w:styleId="Heading1".*?</w:style>', styles, _re.S)
        c = _re.search(r'<w:color w:val="([0-9A-Fa-f]{6})"', m.group(0)) if m else None
        self.accent = c.group(1) if c else "018374"
        return els

    def _normalise(self, els):
        flat, carry = [], False
        for x in els:
            if not _dx_is_p(x):
                if carry:
                    flat.append("<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>")
                    carry = False
                flat.append(x)
                continue
            parts, c2 = _dx_split_breaks(x, self.keep_mark)
            if carry:
                parts[0] = _dx_set_pbb(parts[0])
            flat.extend(parts)
            carry = c2
        # a long run of blank lines used to push the next title onto a new
        # page becomes a real page break
        out, i = [], 0
        while i < len(flat):
            x = flat[i]
            if _dx_filler(x) and not _dx_has_pbb(x):
                j = i
                while j < len(flat) and _dx_filler(flat[j]) and not _dx_has_pbb(flat[j]):
                    j += 1
                if j - i >= 12 and j < len(flat) and _dx_is_p(flat[j]):
                    if not _dx_filler(flat[j]):
                        flat[j] = _dx_set_pbb(flat[j])
                    i = j
                    continue
                out.extend(flat[i:j]); i = j
                continue
            out.append(x); i += 1
        return out

    def _parse(self, els):
        toc_a = next((i for i, x in enumerate(els)
                      if "<w:instrText" in x and " TOC " in x), None)
        toc_b = None
        if toc_a is not None:
            toc_b = next((i for i in range(toc_a, len(els))
                          if 'w:fldCharType="end"' in els[i]), toc_a)
        self.toc = (toc_a, toc_b)
        lo = (toc_b if toc_b is not None else -1) + 1

        idx = {}
        for i in range(lo, len(els)):
            x = els[i]
            if not _dx_is_p(x):
                continue
            t = _dx_text(x).strip()
            m = _re.match(r"Section\s*(\d+)\b", t)
            if m and len(t) < 120:
                idx.setdefault(int(m.group(1)), []).append(i)
        order = sorted(idx, key=lambda n: idx[n][0])

        def start_of(i, floor):
            j = i
            while j - 1 > floor and _dx_filler(els[j - 1]):
                j -= 1
                if _dx_has_pbb(els[j]):
                    return j
            return i

        marks = []
        for n in order:
            hits = idx[n]
            div = hits[0] if len(hits) > 1 else None
            h1 = hits[1] if len(hits) > 1 else hits[0]
            marks.append([n, div, h1])
        secs, prev_end = {}, lo - 1
        bounds = []
        for k, (n, div, h1) in enumerate(marks):
            first = div if div is not None else h1
            bounds.append(start_of(first, prev_end))
            prev_end = h1
        self.front = els[:bounds[0]] if bounds else els
        for k, (n, div, h1) in enumerate(marks):
            a = bounds[k]
            b = bounds[k + 1] if k + 1 < len(bounds) else len(els)
            h1s = start_of(h1, div if div is not None else a - 1) \
                if div is not None else a
            anchors = []
            for i in range(h1 + 1, b):
                x = els[i]
                if not _dx_is_p(x):
                    continue
                t = _dx_text(x).strip()
                if not t or len(t) > 110:
                    continue
                lit = _re.match(r"(\d+)\.(\d+)\.?\s*(.*)$", t)
                styled = _dx_style(x) in self.h2
                if styled or (lit and int(lit.group(1)) == n and len(t) <= 80
                              and not _dx_has_numpr(x)):
                    anchors.append((i, lit.group(3).strip() if lit else t,
                                    bool(lit)))
            starts = []
            floor = h1
            for (i, _t, _l) in anchors:
                starts.append(start_of(i, floor)); floor = i
            subs, auto = [], 0
            for q, (i, title, lit) in enumerate(anchors):
                e = starts[q + 1] if q + 1 < len(anchors) else b
                auto_ref = ""
                if not lit and _dx_has_numpr(els[i]):
                    auto += 1
                    auto_ref = f"{n}.{auto}"
                subs.append(dict(els=els[starts[q]:e], at=i - starts[q],
                                 title=title, literal=lit, sec=n, used=False,
                                 auto_ref=auto_ref))
            head_end = starts[0] if starts else b
            strip = lambda t: _re.sub(r"^Section\s*\d+\s*[-–—:]?\s*", "",
                                      t.strip()).strip()
            title = strip(_dx_text(els[h1]))
            alt = strip(_dx_text(els[div])) if div is not None else title
            toc_div = (div is not None and _dx_style(els[div]) in self.h1
                       and _dx_style(els[h1]) not in self.h1)
            secs[n] = dict(
                num=n, title=title, alt=alt, toc_div=toc_div,
                divider=els[a:h1s] if div is not None else [],
                div_at=(div - a) if div is not None else None,
                head=els[h1s:head_end], h1_at=h1 - h1s, subs=subs)
        self.secs = secs
        self.sec_order = order

    # ── project fields ───────────────────────────────────────────────────
    def _date_label(self):
        raw = (self.d.get("project_date") or "").strip()
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%m/%Y"):
            try:
                return datetime.strptime(raw, fmt).strftime("%m/%Y")
            except ValueError:
                pass
        return raw

    @staticmethod
    def _after_label(x, label_rx, value):
        t = _dx_text_raw(x)
        m = _re.match(label_rx, t)
        if not m:
            return x
        a = m.end()
        while a < len(t) and t[a] in " \u00a0":
            a += 1
        return _dx_sub(x, a, len(t.rstrip(" \u00a0")) if len(t.rstrip(" \u00a0")) > a
                       else len(t), value)

    def _fill_fields(self, els):
        d = self.d
        project = (d.get("project_name") or d.get("client_name") or "").strip()
        customer = (d.get("client_name") or "").strip()
        if customer and (d.get("client_loc") or "").strip():
            customer = f"{customer} – {d['client_loc'].strip()}"
        ref = (d.get("opp_id") or "").strip()
        rules = [(r"\s*Project\s*:", project),
                 (r"\s*Customer\s*:", customer),
                 (r"\s*Proposal Reference Number\s*:", ref),
                 (r"\s*Date\s*:", self._date_label())]
        stop = self.toc[0] if self.toc[0] is not None else min(len(els), 40)
        for i in range(stop):
            x = els[i]
            if _dx_is_p(x):
                for rx, val in rules:
                    if _re.match(rx, _dx_text_raw(x)):
                        els[i] = self._after_label(x, rx, val)
                        break
            elif x.startswith("<w:tbl") and "Contact" in _dx_all_text(x):
                els[i] = self._contacts(x)

        inner_p = _re.compile(r"<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>", _re.S)

        def head_foot(xml):
            def fix(m):
                p = m.group(0)
                t = _dx_text_raw(p)
                if _re.match(r"\s*FOR\s*:", t):
                    return self._after_label(p, r"\s*FOR\s*:", project.upper())
                r = _re.match(r"\s*(\S*)\s*REF\s*:\s*(\S*)\s*$", t)
                if r:
                    mail = (d.get("ae_email") or "").strip() or r.group(2)
                    if r.group(2):
                        p = _dx_sub(p, r.start(2), r.end(2), mail)
                    if r.group(1):
                        p = _dx_sub(p, r.start(1), r.end(1), ref)
                return p
            return inner_p.sub(fix, xml)

        for name in self.names:
            if _re.match(r"word/(header|footer)\d*\.xml$", name):
                self.files[name] = head_foot(
                    self.files[name].decode("utf-8")).encode("utf-8")
        units = _re.sub(r"\D", "", str(d.get("nb_units") or "")) or "1"
        for i, x in enumerate(els):
            if _dx_is_p(x) and "supplied" in x:
                t = _dx_text_raw(x)
                m = _re.match(r"\s*(\d+)\s+lines?\s+(?:is|are)\b", t)
                if m:
                    els[i] = _dx_sub(x, m.start(1), m.end(),
                                     f"{units} line is" if units == "1"
                                     else f"{units} lines are")
        if not project:
            self.report.append(("0.1", "Cover page",
                                "project name empty — cover left blank there"))

    def _advisor(self, tbl):
        """The sales advisor sits in a text box of the contact table. Word
        stores a text box twice (drawing and fallback): both are filled."""
        name = (self.d.get("adv_name") or "").strip()
        mail = (self.d.get("adv_email") or "").strip()
        if not (name or mail):
            return tbl
        par = _re.compile(r"<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>", _re.S)

        def fill(m):
            box = m.group(0)
            live = [(q.start(), q.end(), _dx_text_raw(q.group(0)))
                    for q in par.finditer(box)]
            live = [e for e in live if e[2].strip()]
            k = next((i for i, e in enumerate(live)
                      if "advisor" in e[2].lower()), None)
            if k is None:
                return box
            edits = []
            title = (self.d.get("adv_title") or "").strip()
            if title:
                edits.append((live[k], title))
            if name and k > 0 and "@" not in live[k - 1][2]:
                edits.append((live[k - 1], name))
            hit = next((e for e in live[k + 1:k + 3] if "@" in e[2]), None)
            if mail and hit:
                edits.append((hit, mail))
            for (a, b, text), value in sorted(edits, key=lambda e: -e[0][0]):
                lead = len(text) - len(text.lstrip(" \u00a0"))
                box = box[:a] + _dx_sub(box[a:b], lead, len(text.rstrip()),
                                        value) + box[b:]
            return box

        return _re.sub(r"<w:txbxContent>.*?</w:txbxContent>", fill, tbl,
                       flags=_re.S)

    def _contacts(self, tbl):
        """Application engineer and first sales contact, when the tool has
        them. A field left empty in the tool keeps the skeleton value."""
        d = self.d
        tbl = self._advisor(tbl)
        paras = [(a, b, _dx_text_raw(tbl[a:b]))
                 for a, b in _dx_spans(tbl, ("w:p",))]
        live = [e for e in paras if e[2].strip()]
        edits = {}

        def put(entry, value):
            a, b, t = entry
            if not value:
                return
            lead = len(t) - len(t.lstrip(" \u00a0"))
            edits[a] = (b, _dx_sub(tbl[a:b], lead, len(t.rstrip()), value))

        def around(k, name, mail, title=""):
            if k > 0 and "@" not in live[k - 1][2] \
                    and "contact" not in live[k - 1][2].lower():
                put(live[k - 1], name)
            hit = next((e for e in live[k + 1:k + 3] if "@" in e[2]), None)
            if hit:
                put(hit, mail)
            put(live[k], title)

        ae = next((k for k, e in enumerate(live)
                   if "application engineer" in e[2].lower()), None)
        if ae is not None:
            around(ae, (d.get("ae_name") or "").strip(),
                   (d.get("ae_email") or "").strip(),
                   (d.get("ae_title") or "").strip())
        sales = next((k for k, e in enumerate(live)
                      if "sales" in e[2].lower() and k != ae
                      and "advisor" not in e[2].lower()
                      and "contact" not in e[2].lower()), None)
        if sales is not None:
            around(sales, (d.get("sales_name") or "").strip(),
                   (d.get("sales_email") or "").strip(),
                   (d.get("sales_title") or "").strip())
        out, cur = [], 0
        for a in sorted(edits):
            b, new = edits[a]
            out.append(tbl[cur:a]); out.append(new); cur = b
        return "".join(out) + tbl[cur:]

    # ── matching the tool structure to the skeleton ─────────────────────
    def _match_section(self, sec, taken):
        best, score = None, 0.0
        for n, ws in self.secs.items():
            if n in taken:
                continue
            s = max(_title_sim(sec["title"], ws["title"]),
                    _title_sim(sec["title"], ws["alt"]))
            if n == sec["num"]:
                s += 0.05
            if s > score:
                best, score = n, s
        return best if score >= 0.5 else None

    def _match_sub(self, item, wnum):
        best, score = None, 0.0
        for n, ws in self.secs.items():
            for sub in ws["subs"]:
                if sub["used"]:
                    continue
                s = max(_title_sim(item["title"], sub["title"]),
                        _title_sim(self.title_of(item), sub["title"]))
                if n == wnum:
                    s += 0.05
                if s > score:
                    best, score = sub, s
        return best if score >= 0.5 else None

    @staticmethod
    def _written(block, anchor_at):
        body = [x for k, x in enumerate(block) if k != anchor_at]
        if any(tag in x for x in body
               for tag in ("<w:drawing", "<w:pict", "<w:tbl")):
            return True
        return len("".join(_dx_all_text(x) for x in body).strip()) > 120

    # ── documents ────────────────────────────────────────────────────────
    def _legacy_pages(self, kind):
        """Scope of supply and commercial conditions come from the form."""
        if not HAS_REPORTLAB:
            return None
        if self._legacy is None:
            self._legacy = ProposalPDF(self.d, self.struct, self.status,
                                       options=dict(self.o, stamp=False),
                                       logo=self.logo)
        out = os.path.join(self.tmp, f"tool_{kind}.pdf")
        if os.path.exists(out):
            return out
        c = rl_canvas.Canvas(out, pagesize=A4)
        if kind == "scope":
            self._legacy._scope_pages(c)
        else:
            self._legacy._commercial_page(c)
        c.save()
        return out

    def _scope_table(self):
        """Scope of supply as a Word table, in the fonts of the skeleton:
        item with its description under it, one mark per line in the column
        of who supplies it. The Option column only exists when used; the AE
        note never leaves the tool."""
        rows = self.d.get("scope_of_supply") or []
        if not rows:
            return ""
        has_opt = any(r.get("resp") == "opt" for r in rows)
        if has_opt:
            widths = (6560, 1300, 1300, 1300)
            heads = ("Item", "Baker Hughes", "Customer", "Option")
            marks = ("bh", "cust", "opt")
        else:
            widths = (7560, 1450, 1450)
            heads = ("Item", "Baker Hughes", "Customer")
            marks = ("bh", "cust")

        def run(text, bold=False, colour=None, size=18):
            rpr = ("<w:b/>" if bold else "") \
                + (f'<w:color w:val="{colour}"/>' if colour else "") \
                + f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
            return ("<w:r><w:rPr>" + rpr + '</w:rPr><w:t xml:space="preserve">'
                    + _html.escape(str(text), quote=False) + "</w:t></w:r>")

        def para(runs, centre=False, before=40, after=40):
            return ('<w:p><w:pPr><w:spacing w:before="%d" w:after="%d"/>'
                    % (before, after)
                    + ('<w:jc w:val="center"/>' if centre else "")
                    + "</w:pPr>" + runs + "</w:p>")

        def cell(w, paras, fill=None, span=0):
            tc = f'<w:tcW w:w="{w}" w:type="dxa"/>'
            if span:
                tc += f'<w:gridSpan w:val="{span}"/>'
            if fill:
                tc += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
            tc += '<w:vAlign w:val="center"/>'
            return "<w:tc><w:tcPr>" + tc + "</w:tcPr>" + paras + "</w:tc>"

        edge = '<w:%s w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        out = ['<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'
               + "".join(edge % k for k in ("top", "left", "bottom", "right",
                                            "insideH", "insideV"))
               + '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'
               + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
               + "</w:tblGrid>",
               "<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>"
               + "".join(cell(w, para(run(h, bold=True, colour="FFFFFF"),
                                      centre=k > 0), fill=self.accent)
                         for k, (w, h) in enumerate(zip(widths, heads)))
               + "</w:tr>"]
        current = None
        for r in rows:
            if r.get("section") and r["section"] != current:
                current = r["section"]
                out.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>"
                           + cell(sum(widths), para(run(current, bold=True)),
                                  fill="F2F2F2", span=len(widths)) + "</w:tr>")
            item = r.get("item") or r.get("label", "")
            desc = r.get("desc", "")
            resp = r.get("resp") or ("bh" if r.get("bh") else
                                     "cust" if r.get("customer") else "")
            first = para(run(item), after=0 if desc else 40)
            if desc:
                first += para(run(desc, colour="595959", size=16), before=0)
            out.append("<w:tr><w:trPr><w:cantSplit/></w:trPr>"
                       + cell(widths[0], first)
                       + "".join(cell(w, para(run("X" if resp == m else "",
                                                  bold=True), centre=True))
                                 for w, m in zip(widths[1:], marks))
                       + "</w:tr>")
        out.append("</w:tbl>")
        return "".join(out)

    def _plan_scope(self):
        """The scope table belongs in the skeleton's own 'Scope of Supply'
        sub-section, in place of the object pasted there by hand."""
        self._scope_sub = None
        if not self.o["scope"] or not (self.d.get("scope_of_supply") or []):
            return
        if not any(it["code"].startswith("2.1") for sec in self.struct
                   for it, _d in _leaves(sec)):
            return
        best, score = None, 0.0
        for ws in self.secs.values():
            for sub in ws["subs"]:
                s = _title_sim("Scope of Supply", sub["title"])
                if s > score:
                    best, score = sub, s
        if best is not None and score >= 0.6:
            self._scope_sub = best

    def _put_scope(self, block):
        table = self._scope_table()
        if not table:
            return block
        out, done = [], False
        for x in block:
            if "<w:object" in x and not done:
                out.append(table); done = True
            else:
                out.append(x)
        if not done:
            while out and _dx_filler(out[-1]) and not _dx_has_pbb(out[-1]):
                out.pop()
            out.append(table)
        out.append("<w:p/>")
        return out

    def _docs_for(self, item, written):
        """PDF files to place under this item."""
        st = self.status.get(item["code"], {})
        state = st.get("state")
        ref, title = item["ref"], self.title_of(item)
        if state == "tool":
            code = item["code"]
            if code.startswith("2.1") and self.o["scope"] \
                    and not getattr(self, "_scope_done", False) \
                    and not item["sub"]:
                self._scope_done = True
                if self._scope_sub is not None:
                    self.report.append((ref, title, "generated — table in the Word skeleton"))
                    return []
                self.report.append((ref, title, "generated from the form"))
                return [self._legacy_pages("scope")]
            if code.startswith("11") and self.o["commercial"] \
                    and not getattr(self, "_comm_done", False):
                self._comm_done = True
                self.report.append((ref, title, "generated from the form"))
                return [self._legacy_pages("commercial")]
            return []
        pdfs = [p for p in (st.get("pdfs") or []) if p]
        names = ", ".join(os.path.basename(p) for p in (st.get("paths") or [])[:3])
        if state == "db" and written and not st.get("attached") \
                and not st.get("lib"):
            self.report.append((ref, title, "written in the Word skeleton"))
            return []
        if pdfs:
            how = names or "inserted"
            if st.get("lib_note"):
                how += f"  ({st['lib_note']})"
            if st.get("nda_note"):
                how += f"  ({st['nda_note']})"
            self.report.append((ref, title, how))
            return pdfs
        if st.get("paths"):
            self.report.append((ref, title, "NOT INSERTED — "
                                + (st.get("conv_note") or "cannot be converted")))
        elif state in ("need", "missing") and not written:
            self.report.append((ref, title, "MISSING — heading page only"))
        elif written:
            self.report.append((ref, title, "written in the Word skeleton"))
        return []

    # ── assembly ─────────────────────────────────────────────────────────
    def _marker(self, text):
        return ('<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr>'
                '<w:color w:val="FFFFFF"/><w:sz w:val="2"/></w:rPr><w:t>'
                + text + '</w:t></w:r></w:p>')

    def _emit(self, block):
        if not block:
            return
        block = list(block)
        if self._need_pbb and _dx_is_p(block[0]):
            block[0] = _dx_set_pbb(block[0])
        if _dx_is_p(block[0]) and _dx_has_pbb(block[0]):
            while self.body and _dx_filler(self.body[-1]) \
                    and not _dx_has_pbb(self.body[-1]):
                self.body.pop()
        self._need_pbb = False
        self.body.extend(block)

    def _slot(self, pdfs):
        if not pdfs:
            return
        key = f"{self.SLOT}{len(self._slots) + 1:04d}"
        self._slots[key] = pdfs
        while self.body and _dx_filler(self.body[-1]) \
                and not _dx_has_pbb(self.body[-1]):
            self.body.pop()
        self.body.append(self._marker(key))
        self._need_pbb = True

    def _sec_mark(self, label):
        key = f"{self.SECM}{len(self._secs) + 1:04d}"
        self._secs[key] = label
        while self.body and _dx_filler(self.body[-1]) \
                and not _dx_has_pbb(self.body[-1]):
            self.body.pop()
        self.body.append(self._marker(key))
        self._need_pbb = True

    def _renum_section(self, block, at, disp):
        if at is None or at >= len(block):
            return block
        block = list(block)
        x = block[at]
        t = _dx_text_raw(x)
        m = _re.search(r"Section\s*(\d+)", t)
        if m:
            block[at] = _dx_sub(x, m.start(1), m.end(1), str(disp))
        return block

    @staticmethod
    def _renum_par(x, ref, title=None):
        """Give a heading (or its contents line) its number and, when the
        title follows the configuration, its title. Untouched when nothing
        changes."""
        numbered = ref not in ("", "—", "-")
        t = _dx_text_raw(x)
        m = _re.match(r"\s*(\d+\.\d+(?:\.\d+)*\.?)(\s*)", t)
        has_tab = "<w:tab/>" in _re.sub(r"<w:pPr>.*?</w:pPr>", "", x, 1, _re.S)
        if m:
            old = m.group(1).rstrip(".")
            if title is not None and t[m.end():].strip() != title.strip():
                x = _dx_sub(x, m.end(), len(t), title)
            if not numbered:
                return _dx_sub(x, m.start(1), m.end(), "")
            if old != ref:
                x = _dx_sub(x, m.start(1), m.end(1), ref)
                if not m.group(2) and (not has_tab or len(ref) > len(old)):
                    k = m.start(1) + len(ref)
                    x = _dx_sub(x, k, k, " ")
            return x
        if title is not None and t.strip() != title.strip():
            x = _dx_sub(x, 0, len(t), title)
        if _dx_has_numpr(x):
            x = _dx_drop_numpr(x)
        if numbered:
            x = _dx_sub(x, 0, 0, ref + " ")
        return x

    def _renum_sub(self, sub, ref, title=None):
        block = list(sub["els"])
        at = sub["at"]
        if title is None and sub["auto_ref"] and sub["auto_ref"] == ref:
            return block                   # Word numbers it the same way
        block[at] = self._renum_par(block[at], ref, title)
        return block

    def _templates(self):
        div = h1 = sub = None
        for n in self.sec_order:
            ws = self.secs[n]
            if div is None and ws["divider"] and not any(
                    t in "".join(ws["divider"]) for t in ("<w:drawing", "<w:pict")):
                div = (ws["divider"], ws["div_at"])
            if h1 is None and ws["head"] and not any(
                    t in ws["head"][ws["h1_at"]] for t in ("<w:drawing", "<w:pict")):
                h1 = ws["head"][ws["h1_at"]]
            for s in ws["subs"]:
                if sub is None and s["literal"] and not any(
                        t in s["els"][s["at"]] for t in ("<w:drawing", "<w:pict")):
                    sub = s["els"][s["at"]]
        return div, h1, sub

    def _with_offset(self, divider, at):
        """A divider title sits lower on its page: blank lines in its own
        style when the skeleton relied on overflow to place it."""
        if at is None or at > 0 or not divider:
            return divider, at
        x = divider[0]
        m = _P_OPEN.match(_dx_open(x))
        sp = _dx_ppr(_dx_open(x))
        ppr = _dx_open(x)[sp[0]:sp[1]] if sp else ""
        blank = _dx_drop_numpr(_dx_drop_pbb(
            _dx_fresh(m.group(0)) + ppr + "</w:p>"))
        pad = [_dx_set_pbb(blank)] + [blank] * (self.OFFSET_LINES - 1)
        return pad + [_dx_drop_pbb(x)] + list(divider[1:]), len(pad)

    def _assemble(self):
        self.body, self._need_pbb = list(self.front), False
        self.toc_entries = []
        t_div, t_h1, t_sub = self._templates()
        self._plan_scope()
        taken, plan = set(), []
        for sec in self.struct:
            if sec["num"] == 0:
                continue
            w = self._match_section(sec, taken)
            if w is not None:
                taken.add(w)
            plan.append((sec, w))
        for n in self.sec_order:
            if n not in taken:
                self.report.append((f"S{n}", self.secs[n]["title"],
                                    "skeleton section not in this offer — left out"))

        def squeeze(text):
            return _re.sub(r"\s+", " ", text).strip()

        for sec, w in plan:
            disp = sec.get("disp", sec["num"])
            ws = self.secs.get(w) if w is not None else None
            head_written = bool(ws) and self._written(ws["head"], ws["h1_at"])
            rows, any_content = [], head_written
            for it, depth in _leaves(sec):
                top = it["ref"] == str(disp)
                sub = None if top else self._match_sub(it, w)
                if sub is not None:
                    sub["used"] = True
                written = (sub is self._scope_sub or
                           self._written(sub["els"], sub["at"])) if sub \
                    else head_written if top else False
                pdfs = self._docs_for(it, written)
                if pdfs or written:
                    any_content = True
                rows.append((it, top, sub, pdfs))
            if self.o["skip_empty"] and not any_content:
                self.report.append((f"S{disp}", sec["title"], "empty — left out"))
                continue
            if ws is None and not any(r[3] for r in rows) \
                    and not self.o["placeholders"]:
                continue

            self._sec_mark(f"{int(disp):02d}_{sec['title']}"
                           if str(disp).isdigit() else f"{disp}_{sec['title']}")
            if ws is not None:
                divider, at = self._with_offset(ws["divider"], ws["div_at"])
                divider = self._renum_section(divider, at, disp)
                self._emit(divider)
                head = self._renum_section(ws["head"], ws["h1_at"], disp)
                self._emit(head)
                src = divider[at] if ws["toc_div"] else head[ws["h1_at"]]
                self.toc_entries.append(dict(level=1, xml=src, disp=disp,
                                             label=squeeze(_dx_text(src))))
            else:
                if t_div:
                    blk, at = list(map(_dx_fresh, t_div[0])), t_div[1]
                    blk[at] = _dx_set_text(blk[at], f"Section {disp} - {sec['title']}")
                    blk, at = self._with_offset(blk, at)
                    blk[0] = _dx_set_pbb(blk[0])
                    self._emit(blk)
                label = f"Section {disp} {sec['title']}"
                if t_h1:
                    self._emit([_dx_set_pbb(_dx_set_text(_dx_fresh(t_h1), label))])
                self.toc_entries.append(dict(level=1, xml="", disp=disp,
                                             label=label))

            late = []
            inline_first = bool(ws and ws["subs"]
                                and not _dx_has_pbb(ws["subs"][0]["els"][0]))
            for it, top, sub, pdfs in rows:
                if top:
                    late.append(pdfs)
                    continue
                ref = it["ref"]
                numbered = ref not in ("", "—", "-")
                if sub is not None and not pdfs and not self._written(sub["els"], sub["at"]) \
                        and sub is not self._scope_sub and not self.o["placeholders"]:
                    self.report.append((ref, sub["title"],
                                        "no document — heading left out"))
                    continue
                if sub is not None:
                    title = self.title_of(it) if it.get("tdyn") else None
                    block = self._renum_sub(sub, ref, title)
                    if not pdfs and sub is not self._scope_sub \
                            and not self._written(sub["els"], sub["at"]):
                        block = list(block[:sub["at"] + 1]) + [
                            '<w:p><w:r><w:rPr><w:i/><w:color w:val="595959"/></w:rPr>'
                            '<w:t xml:space="preserve">This document will be provided '
                            'with the next revision of the proposal.</w:t></w:r></w:p>']
                    if pdfs and self.status.get(it["code"], {}).get("replace"):
                        # the concept document replaces the generic text
                        block = block[:sub["at"] + 1]
                    if sub is self._scope_sub:
                        block = self._put_scope(block)
                    opens = self.toc_entries[-1]["level"] == 1 and not self._need_pbb
                    if opens and inline_first:
                        block[0] = _dx_drop_pbb(block[0])
                    elif not opens and not _dx_has_pbb(block[0]):
                        block[0] = _dx_set_pbb(block[0])
                    self._emit(block)
                    shown = squeeze(_dx_text(block[sub["at"]]))
                    self.toc_entries.append(dict(
                        level=2, xml=block[sub["at"]], ref=ref, title=title,
                        plain=_re.sub(r"^\d+(?:\.\d+)*\.?\s*", "", shown),
                        label=shown))
                    self._slot(pdfs)
                elif pdfs or (self.o["placeholders"] and not it["sub"]
                              and (ws is None or it.get("custom"))):
                    plain = self.title_of(it)
                    label = f"{ref} {plain}" if numbered else plain
                    if t_sub:
                        opens = (self.toc_entries and self.toc_entries[-1]["level"] == 1
                                 and not self._need_pbb)
                        hp = _dx_drop_numpr(_dx_set_text(_dx_fresh(t_sub), label))
                        self._emit([_dx_drop_pbb(hp) if opens else _dx_set_pbb(hp)])
                    self.toc_entries.append(dict(level=2, xml="", ref=ref,
                                                 title=plain, plain=plain,
                                                 label=label))
                    self._slot(pdfs)
            if ws is not None:
                for sub in ws["subs"]:
                    if not sub["used"] and sub["sec"] == w:
                        self.report.append(
                            ("—", sub["title"],
                             "skeleton sub-section not in this offer — left out"))
            for pdfs in late:
                self._slot(pdfs)
        while self.body and _dx_filler(self.body[-1]) \
                and not _dx_has_pbb(self.body[-1]) and not self._need_pbb:
            self.body.pop()
        # one empty page of the skeleton, header and footer only: the frame
        # put behind the documents that are inserted
        self._frame = bool(self.o["frame"] and self._slots)
        if self._frame:
            self.body.append("<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>")

    def _rebuild_toc(self):
        """Contents in the new order. Each line is the skeleton's own line
        for that heading (tabs, link and character style included), with
        only its number brought up to date; a heading the skeleton did not
        list gets a copy of a neighbouring line."""
        a, b = self.toc
        if a is None or not self.toc_entries:
            return
        toc = self.body[a:b + 1]
        run = r"<w:r(?:\s[^>]*)?>(?:(?!</w:r>).)*?w:fldCharType=\"%s\"/>(?:(?!</w:r>).)*?</w:r>"
        first = toc[0]
        sp = _dx_ppr(first)
        cut0 = sp[1] if sp else _P_OPEN.match(first).end()
        sep = _re.search(run % "separate", first, _re.S)
        prefix = first[cut0:sep.end()] if sep else ""
        if sep:
            toc[0] = first[:cut0] + first[sep.end():]
        end = _re.search(run % "end", toc[-1], _re.S)
        closing = None
        if end:
            bare = toc[-1][:end.start()] + toc[-1][end.end():]
            if _dx_all_text(bare).strip():
                toc[-1] = bare
            else:
                closing = toc.pop()

        def anchor(x):
            m = _re.search(r'<w:hyperlink [^>]*w:anchor="([^"]+)"', x)
            return m.group(1) if m else ""

        def level(x):
            st = _dx_style(x)
            return 1 if st.endswith("1") else 2 if st.endswith("2") else 0

        by_anchor = {anchor(x): x for x in toc if anchor(x)}
        front_xml = "".join(self.body[:a]) + "".join(self.body[b + 1:len(self.front)])
        lines = [x for x in toc
                 if anchor(x) and f'w:name="{anchor(x)}"' in front_xml]
        unlink = lambda x: _re.sub(r"</?w:hyperlink\b[^>]*>", "", _dx_fresh(x))
        num_rx = _re.compile(r"\s*\d+\.\d+")
        t1 = next((x for x in toc if level(x) == 1 and x not in lines
                   and "<w:tab/>" not in x[x.find("</w:pPr>"):]), None) \
            or next((x for x in toc if level(x) == 1 and x not in lines), None)
        t2n = next((x for x in toc if level(x) == 2
                    and num_rx.match(_dx_text_raw(x))), None)
        t2 = next((x for x in toc if level(x) == 2 and x not in lines
                   and not num_rx.match(_dx_text_raw(x))), None)

        for e in self.toc_entries:
            names = _re.findall(r'w:name="([^"]+)"', e["xml"] or "")
            orig = next((by_anchor[n] for n in names if n in by_anchor), None)
            if e["level"] == 1:
                if orig is not None:
                    x = self._renum_section([orig], 0, e["disp"])[0]
                elif t1 is not None:
                    x = unlink(t1)
                    x = _dx_sub(x, 0, len(_dx_text_raw(x)), e["label"])
                else:
                    continue
            else:
                numbered = e["ref"] not in ("", "—", "-")
                if orig is not None:
                    x = self._renum_par(orig, e["ref"], e.get("title"))
                elif numbered and t2n is not None:
                    x = self._renum_par(unlink(t2n), e["ref"], e["plain"])
                elif t2 is not None or t2n is not None:
                    x = unlink(t2 if t2 is not None else t2n)
                    x = _dx_sub(x, 0, len(_dx_text_raw(x)), e["label"])
                    if t2 is None:
                        x = x.replace("<w:tab/>", "")
                else:
                    continue
            lines.append(x)
        if not lines:
            return
        if prefix:
            sp = _dx_ppr(lines[0])
            k = sp[1] if sp else _P_OPEN.match(lines[0]).end()
            lines[0] = lines[0][:k] + prefix + lines[0][k:]
        if closing is not None:
            lines.append(closing)
        elif end:
            lines[-1] = lines[-1][:-len("</w:p>")] + end.group(0) + "</w:p>"
        self.body[a:b + 1] = lines

    # ── output ───────────────────────────────────────────────────────────
    def _write_docx(self, path, clean=False):
        body = self.body
        if clean:
            # the copy kept for the user: no placeholder page, the documents
            # themselves only exist in the PDF
            body = [x for x in body
                    if not _re.search(f"(?:{self.SLOT}|{self.SECM})\\d{{4}}", x)]
            if self._frame and body:
                body = body[:-1]
        doc = self.pre + "".join(body + self.tail) + self.post
        self.files["word/document.xml"] = doc.encode("utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.names:
                z.writestr(n, self.files[n])

    def build(self, out_pdf, keep_docx=None):
        if not HAS_PYPDF:
            raise RuntimeError("pypdf is required.")
        els = self._normalise(self._load())
        self._parse(els)
        if not self.secs:
            raise RuntimeError("No 'Section n' heading found in the Word skeleton.")
        self._fill_fields(els)
        self._parse(els)                    # blocks cut again, fields filled
        self._assemble()
        self._rebuild_toc()
        docx = os.path.join(self.tmp, "proposal_build.docx")
        self._write_docx(docx)
        if keep_docx:
            try:
                self._write_docx(keep_docx, clean=True)
            except OSError:
                pass                        # open in Word: the PDF still comes
        raw = os.path.join(self.tmp, "proposal_build.pdf")
        self.engine = _office_to_pdf(docx, raw)
        if not self.engine:
            raise RuntimeError(
                "The Word file was filled, but it could not be turned into "
                "a PDF on this computer.\n\n" + "\n".join(CONVERSION_LOG))

        reader = PdfReader(raw)
        writer = PdfWriter()
        pages = list(reader.pages)
        frame = pages.pop() if self._frame and len(pages) > 1 else None
        try:
            from pypdf import PageObject, Transformation
        except Exception:
            frame = None

        mg_ = self.margins

        def body_lines(page):
            """(lowest y, lines) of the text between header and footer."""
            ph_ = float(page.mediabox.height)
            top_, bot_ = ph_ - mg_["top"] + 2, mg_["bottom"] - 2
            got = []

            def v(text, cm, tm, _f, _s):
                if not text or not text.strip():
                    return
                y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
                if bot_ < y < top_:
                    got.append((round(y, 1), text.strip()))
            try:
                page.extract_text(visitor_text=v)
            except Exception:
                return None, []
            if not got:
                return None, []
            lines = {}
            for y, t in got:
                lines.setdefault(y, []).append(t)
            ordered = [" ".join(lines[y]) for y in sorted(lines, reverse=True)]
            return min(y for y, _ in got), ordered

        def onto_host(host, q):
            """Put the first page of a document under the heading that
            announces it, when there is room for it at a readable size."""
            try:
                w, h = float(q.mediabox.width), float(q.mediabox.height)
                if w > h * 1.05 or q.get("/Rotate") not in (None, 0):
                    return False
                low, _lines = body_lines(host)
                pw, ph = float(host.mediabox.width), float(host.mediabox.height)
                bw = pw - mg_["left"] - mg_["right"]
                full_h = ph - mg_["top"] - mg_["bottom"] - 12
                top = (low - 14) if low is not None else ph - mg_["top"] - 6
                room = top - mg_["bottom"] - 6
                k_full = min(bw / w, full_h / h, 1.0)
                k = min(bw / w, room / h, 1.0)
                if room <= 0 or k < 0.80 * k_full:
                    return False
                op = (Transformation()
                      .translate(-float(q.mediabox.left), -float(q.mediabox.bottom))
                      .scale(k, k)
                      .translate(mg_["left"] + (bw - w * k) / 2, top - h * k))
                host.merge_transformed_page(q, op)
                return True
            except Exception:
                return False

        def banner(q, title):
            """A drawing after a page that only holds its heading: the
            heading goes on a band above the drawing, the empty page goes."""
            try:
                try:
                    q.transfer_rotation_to_content()
                except Exception:
                    pass
                w, h = float(q.mediabox.width), float(q.mediabox.height)
                band = 26.0
                k = (h - band - 4) / h
                buf = io.BytesIO()
                from reportlab.pdfgen import canvas as _cv
                c = _cv.Canvas(buf, pagesize=(w, h))
                c.setFillColorRGB(*_rl_colour(BH_GREEN))
                c.rect(18, h - band + 6, 3, band - 12, fill=1, stroke=0)
                c.setFillColorRGB(*_rl_colour(BH_INK))
                c.setFont("Helvetica-Bold", 11)
                c.drawString(28, h - band + 10, title[:140])
                c.save()
                buf.seek(0)
                new = PageObject.create_blank_page(width=w, height=h)
                new.merge_page(PdfReader(buf).pages[0])
                op = (Transformation()
                      .translate(-float(q.mediabox.left), -float(q.mediabox.bottom))
                      .scale(k, k).translate((w - w * k) / 2, 2))
                new.merge_transformed_page(q, op)
                return new
            except Exception:
                return None

        def framed(q):
            """An inserted A4 portrait page sits inside the skeleton page,
            under its header and above its footer. Drawings and landscape
            sheets keep their own page and size."""
            if frame is None:
                return q
            try:
                w, h = float(q.mediabox.width), float(q.mediabox.height)
                if w > h * 1.05 or q.get("/Rotate") not in (None, 0):
                    return q
                pw, ph = float(frame.mediabox.width), float(frame.mediabox.height)
                mg, pad = self.margins, 6.0
                bw = pw - mg["left"] - mg["right"]
                bh = ph - mg["top"] - mg["bottom"] - 2 * pad
                k = min(bw / w, bh / h, 1.0)
                new = PageObject.create_blank_page(width=pw, height=ph)
                new.merge_page(frame)
                op = (Transformation()
                      .translate(-float(q.mediabox.left), -float(q.mediabox.bottom))
                      .scale(k, k)
                      .translate(mg["left"] + (bw - w * k) / 2,
                                 ph - mg["top"] - pad - h * k))
                new.merge_transformed_page(q, op)
                return new
            except Exception:
                return q

        self.section_starts = [("00_Front matter", 0)]
        outline = []
        host = None             # last skeleton page written, if it is the last page
        self.joined = 0
        for pg in pages:
            try:
                txt = pg.extract_text() or ""
            except Exception:
                txt = ""
            m = _re.search(self.SLOT + r"\d{4}", txt)
            if m and m.group(0) in self._slots:
                first = True
                for p in self._slots[m.group(0)]:
                    try:
                        for q in PdfReader(p).pages:
                            if first and host is not None:
                                first = False
                                if onto_host(host, q):
                                    self.joined += 1
                                    continue
                                low, lines = body_lines(host)
                                ph_ = float(host.mediabox.height)
                                heading_only = (lines and len(lines) <= 2 and low is not None
                                                and low > ph_ * 0.55)
                                w_, h_ = float(q.mediabox.width), float(q.mediabox.height)
                                rot = q.get("/Rotate") in (90, 270, -90)
                                if heading_only and (w_ > h_ * 1.05) != rot \
                                        and hasattr(writer, "remove_page"):
                                    nb = banner(q, "   ".join(lines))
                                    if nb is not None:
                                        writer.remove_page(len(writer.pages) - 1)
                                        if outline and outline[-1][1] >= len(writer.pages):
                                            outline[-1] = (outline[-1][0], len(writer.pages))
                                        self.section_starts = [
                                            (lb, min(i, len(writer.pages)))
                                            for lb, i in self.section_starts]
                                        writer.add_page(nb)
                                        self.joined += 1
                                        continue
                            first = False
                            writer.add_page(framed(q))
                    except Exception as e:
                        self.report.append(("!", os.path.basename(p),
                                            f"unreadable PDF ({e})"))
                host = None
                continue
            m = _re.search(self.SECM + r"\d{4}", txt)
            if m and m.group(0) in self._secs:
                label = self._secs[m.group(0)]
                self.section_starts.append((label, len(writer.pages)))
                outline.append((label.split("_", 1)[-1], len(writer.pages)))
                continue
            host = writer.add_page(pg)
        for title, page in outline:
            try:
                writer.add_outline_item(title, page)
            except Exception:
                try:
                    writer.addBookmark(title, page)
                except Exception:
                    pass
        with open(out_pdf, "wb") as fh:
            writer.write(fh)
        return len(writer.pages)


# ─────────────────────────────────────────────────────────────────────────────
# WIDGET HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _shade(hex_color, factor=0.86):
    """Darken (factor<1) or lighten (factor>1) a #RRGGBB colour."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    f = lambda c: max(0, min(255, int(c * factor)))
    return f"#{f(r):02X}{f(g):02X}{f(b):02X}"


def _entry(parent, var, w=28, state=tk.NORMAL, fnt=None, **kw):
    """Flat field with an animated focus ring in Baker Hughes green."""
    e = tk.Entry(parent, textvariable=var, width=w, font=fnt or FONT_N,
                 relief=tk.FLAT, bd=0, highlightthickness=1,
                 highlightbackground=BH_BORDER, highlightcolor=BH_GREEN,
                 bg=BH_WHITE, fg=BH_INK, insertbackground=BH_GREEN,
                 disabledbackground=BH_BG, disabledforeground=BH_MUTED,
                 selectbackground=BH_SEC, selectforeground=BH_INK,
                 state=state, **kw)

    def ring(to):
        Motion.run(e, 120,
                   lambda t: e.config(highlightbackground=_mix(
                       e.cget("highlightbackground"), to, t)),
                   key=("ring", id(e)))
    e.bind("<FocusIn>",  lambda _e: ring(BH_GREEN))
    e.bind("<FocusOut>", lambda _e: ring(BH_BORDER))
    e.bind("<Enter>", lambda _e: e.config(highlightbackground=BH_GRID)
           if e.focus_get() is not e else None)
    e.bind("<Leave>", lambda _e: e.config(highlightbackground=BH_BORDER)
           if e.focus_get() is not e else None)
    return e

def _combo(parent, var, vals, w=22, **kw):
    return ttk.Combobox(parent, textvariable=var, values=vals, width=w,
                        font=FONT_N, state="readonly", style="BH.TCombobox", **kw)

def _btn(parent, text, cmd, bg=BH_GREEN, fg=BH_WHITE, fnt=None, w=None,
         pad=(13, 6), **kw):
    """Solid button: eased hover, pressed state, no focus rectangle."""
    fnt = fnt or FONT_B
    hv = _shade(bg, 1.12 if bg in (BH_GRAY, BH_MUTED) else 1.10)
    pr = _shade(bg, 0.88)
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, font=fnt,
                  relief=tk.FLAT, cursor="hand2", activebackground=pr,
                  activeforeground=fg, padx=pad[0], pady=pad[1], bd=0,
                  highlightthickness=0, **kw)
    if w:
        b.config(width=w)
    b.bind("<Enter>", lambda e: _fade([b], bg, hv, 110, key=("b", id(b))))
    b.bind("<Leave>", lambda e: _fade([b], hv, bg, 130, key=("b", id(b))))
    b.bind("<ButtonPress-1>", lambda e: (Motion.stop(("b", id(b))),
                                         b.config(bg=pr)))
    b.bind("<ButtonRelease-1>", lambda e: b.config(bg=hv))
    return b

def _ghost_btn(parent, text, cmd, fg=BH_GREEN, **kw):
    """Flat, low-emphasis button for in-row actions."""
    bg = kw.pop("bg", BH_WHITE)
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg, font=FONT_SM,
                  relief=tk.FLAT, cursor="hand2", bd=0, padx=8, pady=2,
                  activebackground=_shade(BH_SEC, 0.96), activeforeground=fg,
                  highlightthickness=0, **kw)
    b.bind("<Enter>", lambda e: _fade([b], bg, BH_SEC, 100, key=("g", id(b))))
    b.bind("<Leave>", lambda e: _fade([b], BH_SEC, bg, 120, key=("g", id(b))))
    return b

def _lbl(parent, text, fnt=None, fg=BH_INK, bg=BH_WHITE, **kw):
    return tk.Label(parent, text=text, font=fnt or FONT_N, fg=fg, bg=bg, **kw)


def _mix(c1, c2, t):
    a = [int(c1.lstrip("#")[i:i+2], 16) for i in (0, 2, 4)]
    b = [int(c2.lstrip("#")[i:i+2], 16) for i in (0, 2, 4)]
    return "#%02X%02X%02X" % tuple(int(a[i] + (b[i]-a[i])*t) for i in range(3))


def _grad_bar(parent, height=2, c1=BH_GREEN, c2=None, bg=BH_DARK, steps=None):
    """Flat brand rule. Kept as a function so every existing call site gets the
    same restrained hairline instead of a decorative gradient."""
    return tk.Frame(parent, bg=c1, height=max(int(height), 1))


def _round_grad(cv, x0, y0, x1, y1, r, c1=BH_GREEN, c2=None, steps=None):
    """Flat fill with softly rounded corners — no gradient, no shine."""
    _round_rect(cv, x0, y0, x1, y1, min(r, 4), fill=BH_GREEN, outline=BH_GREEN)


def _brand_mark(parent, size=42, bg=BH_DARK):
    """Kept for compatibility: a plain rule, no badge, no monogram."""
    return tk.Frame(parent, bg=BH_GREEN, width=3, height=size)


def _round_rect(cv, x1, y1, x2, y2, r, **kw):
    pts = [x1+r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2,
           x2-r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


_PILL_FONT = {}


def _pill(parent, text, bg, fg=None, host_bg=BH_WHITE, icon="", h=20):
    """Status tag: tinted background, coloured text, square corners.
    The `icon` argument is accepted and ignored on purpose."""
    key = (UI["fam"], max(7, int(8 * UI["scale"])))
    f = _PILL_FONT.get(key)
    if f is None:
        f = _PILL_FONT[key] = tkfont.Font(family=key[0], size=key[1], weight="bold")
    on_dark = _lum(host_bg) < 0.42
    tint = _mix(bg, host_bg, 0.80 if on_dark else 0.86)
    line = _mix(bg, host_bg, 0.55)
    ink = _mix(bg, BH_WHITE, 0.35) if on_dark else _shade(bg, 0.86)
    lb = tk.Label(parent, text=f" {text} ", font=f, fg=fg or ink, bg=tint,
                  bd=0, highlightthickness=1, highlightbackground=line,
                  highlightcolor=line, padx=7,
                  pady=max(2, int(2 * UI["scale"])))
    return lb


# ─────────────────────────────────────────────────────────────────────────────
# DESIGN SYSTEM  ·  motion, anti-aliased primitives, feedback
# Everything below is presentation only: no business rule depends on it.
# ─────────────────────────────────────────────────────────────────────────────
from PIL import ImageDraw, ImageFilter

UI = {"scale": 1.0, "fam": "Segoe UI", "mono": "Consolas", "motion": True}

# Radii and the 4 pt spacing scale used across the shell
R_SM, R_MD, R_LG, R_PILL = 5, 8, 12, 999
SP1, SP2, SP3, SP4, SP5 = 4, 8, 14, 18, 26

BH_CARD_BR  = "#E1E5E3"          # derived — card outline
BH_HOVER    = "#F1F4F3"          # derived — row hover on white
BH_SHADOW_1 = _mix(BH_BG, BH_GRID, 0.60)
BH_SHADOW_2 = _mix(BH_BG, BH_GRID, 0.24)


def _resolve_fonts(root):
    """Pick the best UI family available on this machine and rebuild the type
    scale for the current display DPI. Called once, before any widget exists."""
    global FONT_T, FONT_ST, FONT_S, FONT_B, FONT_N, FONT_SM, FONT_XS, FONT_BD
    global FONT_C, FONT_CS
    try:
        fams = set(tkfont.families(root))
    except Exception:
        fams = set()
    for cand in ("Segoe UI Variable Text", "Segoe UI", "Inter", "Roboto",
                 "Noto Sans", "DejaVu Sans", "Helvetica"):
        if cand in fams:
            UI["fam"] = cand
            break
    for cand in ("Cascadia Mono", "Consolas", "JetBrains Mono",
                 "DejaVu Sans Mono", "Courier New"):
        if cand in fams:
            UI["mono"] = cand
            break
    try:
        UI["scale"] = max(1.0, min(1.5, root.winfo_fpixels("1i") / 96.0))
    except Exception:
        UI["scale"] = 1.0
    f, m, k = UI["fam"], UI["mono"], UI["scale"]
    z = lambda n: max(7, int(round(n * k)))
    FONT_T  = (f, z(22), "bold")
    FONT_ST = (f, z(12), "bold")
    FONT_S  = (f, z(10), "bold")
    FONT_B  = (f, z(9),  "bold")
    FONT_N  = (f, z(9))
    FONT_SM = (f, z(8))
    FONT_XS = (f, z(7))
    FONT_BD = (f, z(8),  "bold")
    FONT_C  = (m, z(9),  "bold")
    FONT_CS = (m, z(8))


def _fnt(size, bold=False):
    """Ad-hoc font on the resolved family, DPI-scaled."""
    s = max(7, int(round(size * UI["scale"])))
    return (UI["fam"], s, "bold") if bold else (UI["fam"], s)


def _rgb(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i+2], 16) for i in (0, 2, 4))


def _lum(c):
    r, g, b = (v / 255.0 for v in _rgb(c))
    f = lambda v: v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _on(bg_colour):
    """Readable text colour for any background — WCAG-style contrast pick."""
    return BH_WHITE if _lum(bg_colour) < 0.42 else BH_INK


# ── motion engine ────────────────────────────────────────────────────────────
class Motion:
    """Frame-based tweening. One generation counter per key, so a new
    animation silently supersedes the one it replaces instead of fighting it."""
    MS = 16
    _gen = {}

    @staticmethod
    def ease_out(t):
        return 1 - (1 - t) ** 3

    @staticmethod
    def ease_in_out(t):
        return 4 * t * t * t if t < 0.5 else 1 - ((-2 * t + 2) ** 3) / 2

    @staticmethod
    def linear(t):
        return t

    @classmethod
    def run(cls, widget, ms, step, key=None, ease=None, done=None):
        ease = ease or cls.ease_out
        key = key if key is not None else id(widget)
        if not UI.get("motion", True):
            try:
                step(1.0)
                if done:
                    done()
            except Exception:
                pass
            return
        gen = cls._gen.get(key, 0) + 1
        cls._gen[key] = gen
        n = max(int(ms / cls.MS), 1)

        def tick(i=0):
            if cls._gen.get(key) != gen:
                return
            i += 1
            t = min(i / n, 1.0)
            try:
                step(ease(t))
            except tk.TclError:
                return
            except Exception:
                return
            if t < 1.0:
                try:
                    widget.after(cls.MS, lambda: tick(i))
                except tk.TclError:
                    return
            elif done:
                try:
                    done()
                except Exception:
                    pass
        try:
            widget.after(cls.MS, tick)
        except tk.TclError:
            pass

    @classmethod
    def stop(cls, key):
        cls._gen[key] = cls._gen.get(key, 0) + 1


def _fade(widgets, frm, to, ms=130, key=None):
    """Animate the background of a group of widgets from one colour to another."""
    ws = [w for w in widgets if w is not None]
    if not ws:
        return
    def step(t):
        c = _mix(frm, to, t)
        for w in ws:
            w.config(bg=c)
    Motion.run(ws[0], ms, step, key=key or ("fade", id(ws[0])))


def _fade_fg(widget, frm, to, ms=130):
    Motion.run(widget, ms, lambda t: widget.config(fg=_mix(frm, to, t)),
               key=("fg", id(widget)))


# ── anti-aliased primitives (PIL renders the shape, Tk draws the text) ───────
_IMG_CACHE = {}


def _grad_img(w, h, c1, c2, horizontal=True):
    n = max(w if horizontal else h, 2)
    a, b = _rgb(c1), _rgb(c2)
    px = [tuple(int(a[k] + (b[k] - a[k]) * i / (n - 1)) for k in range(3))
          for i in range(n)]
    strip = Image.new("RGB", (n, 1) if horizontal else (1, n))
    strip.putdata(px)
    return strip.resize((max(w, 1), max(h, 1)), Image.Resampling.BILINEAR)


def _round_img(w, h, radius, fill=None, bg=BH_WHITE, border=None, grad=None,
               shadow=0, key=None, ss=3):
    """Anti-aliased rounded rectangle flattened onto `bg`, returned as a
    PhotoImage. `grad` is (colour_a, colour_b[, horizontal])."""
    w, h = max(int(w), 1), max(int(h), 1)
    ck = key or ("rr", w, h, radius, fill, bg, border, grad, shadow)
    hit = _IMG_CACHE.get(ck)
    if hit is not None:
        return hit
    sh_h = max(h - shadow, 1)
    r = max(0, min(radius, min(w, sh_h) / 2))
    base = Image.new("RGB", (w, h), bg)
    mask = Image.new("L", (w * ss, sh_h * ss), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, w * ss - 1, sh_h * ss - 1], radius=r * ss, fill=255)
    mask = mask.resize((w, sh_h), Image.Resampling.LANCZOS)
    full = Image.new("L", (w, h), 0)
    full.paste(mask, (0, 0))
    if shadow:
        blur = full.filter(ImageFilter.GaussianBlur(max(shadow * 0.8, 1)))
        moved = Image.new("L", (w, h), 0)
        moved.paste(blur, (0, shadow))
        moved = moved.point(lambda v: int(v * 0.55))
        base = Image.composite(Image.new("RGB", (w, h), _mix(bg, BH_NAVY, 0.30)),
                               base, moved)
    layer = (_grad_img(w, h, grad[0], grad[1],
                       grad[2] if len(grad) > 2 else True)
             if grad else Image.new("RGB", (w, h), fill or BH_GREEN))
    base = Image.composite(layer, base, full)
    if border:
        bm = Image.new("L", (w * ss, sh_h * ss), 0)
        ImageDraw.Draw(bm).rounded_rectangle(
            [0, 0, w * ss - 1, sh_h * ss - 1], radius=r * ss,
            outline=255, width=max(int(ss), 1))
        bm = bm.resize((w, sh_h), Image.Resampling.LANCZOS)
        bf = Image.new("L", (w, h), 0)
        bf.paste(bm, (0, 0))
        base = Image.composite(Image.new("RGB", (w, h), border), base, bf)
    img = ImageTk.PhotoImage(base)
    _IMG_CACHE[ck] = img
    return img


class _SoftButton(tk.Label):
    """Rounded, anti-aliased, animated button. The shape is a PIL image, the
    label is real Tk text, so it stays crisp at any DPI."""

    def __init__(self, parent, text, command=None, fill=None, grad=None,
                 fg=None, host_bg=BH_WHITE, fnt=None, pad=(20, 10),
                 radius=R_SM, border=None, shadow=0, min_w=0, **kw):
        self._cmd = command
        self._enabled = True
        fnt = fnt or FONT_B
        f = tkfont.Font(family=fnt[0], size=fnt[1],
                        weight=(fnt[2] if len(fnt) > 2 else "normal"))
        w = max(f.measure(text) + pad[0] * 2, min_w)
        h = f.metrics("linespace") + pad[1] * 2 + shadow
        self._base = dict(w=w, h=h, radius=radius, bg=host_bg,
                          border=border, shadow=shadow)
        self._fill, self._grad = fill, grad
        self._imgs = {
            "n": self._img(1.00),
            "h": self._img(1.10),
            "p": self._img(0.90),
            "d": self._img(1.00, dim=True),
        }
        ref = (grad[0] if grad else (fill or BH_GREEN))
        super().__init__(parent, text=text, image=self._imgs["n"],
                         compound=tk.CENTER, font=fnt, bd=0,
                         highlightthickness=0, bg=host_bg,
                         fg=fg or _on(ref), cursor="hand2", **kw)
        self._fg = fg or _on(ref)
        self.bind("<Enter>", lambda e: self._to("h"))
        self.bind("<Leave>", lambda e: self._to("n"))
        self.bind("<ButtonPress-1>", lambda e: self._to("p"))
        self.bind("<ButtonRelease-1>", self._release)

    def _img(self, k, dim=False):
        b = self._base
        tone = lambda c: _mix(_shade(c, k), BH_MUTED, 0.55) if dim else _shade(c, k)
        grad = ((tone(self._grad[0]), tone(self._grad[1])) if self._grad else None)
        return _round_img(b["w"], b["h"], b["radius"],
                          fill=None if grad else tone(self._fill or BH_GREEN),
                          bg=b["bg"], border=b["border"], grad=grad,
                          shadow=b["shadow"],
                          key=("btn", b["w"], b["h"], b["radius"], b["bg"],
                               b["border"], b["shadow"], self._fill,
                               self._grad, round(k, 2), dim))

    def _to(self, state):
        if not self._enabled:
            return
        self.config(image=self._imgs[state])

    def _release(self, e):
        if not self._enabled:
            return
        inside = 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height()
        self._to("h" if inside else "n")
        if inside and self._cmd:
            self._cmd()

    def set_enabled(self, on):
        self._enabled = bool(on)
        self.config(image=self._imgs["n" if on else "d"],
                    cursor="hand2" if on else "arrow",
                    fg=self._fg if on else BH_MUTED)


# ── scrolling ────────────────────────────────────────────────────────────────
class _SmoothScroll:
    """Inertial scrolling for a canvas: the wheel moves a target, a critically
    damped follower chases it. Wheel spam accumulates instead of stuttering."""

    def __init__(self, canvas):
        self.cv = canvas
        self.target = None
        self.running = False

    def _bounds(self):
        bb = self.cv.bbox("all") or (0, 0, 1, 1)
        total = max(bb[3], 1)
        view = max(self.cv.winfo_height(), 1)
        return total, max(total - view, 0)

    def by(self, px):
        total, maxy = self._bounds()
        cur = self.cv.canvasy(0) if self.target is None else self.target
        self.target = max(0.0, min(float(maxy), cur + px))
        if not self.running:
            self._loop()

    def to(self, y):
        total, maxy = self._bounds()
        self.target = max(0.0, min(float(maxy), float(y)))
        if not self.running:
            self._loop()

    def to_widget(self, widget, margin=18):
        try:
            top = widget.winfo_rooty() - self.cv.winfo_rooty() + self.cv.canvasy(0)
        except tk.TclError:
            return
        self.to(top - margin)

    def _loop(self):
        self.running = True

        def tick():
            try:
                total, maxy = self._bounds()
                cur = self.cv.canvasy(0)
                if self.target is None:
                    self.running = False
                    return
                d = self.target - cur
                if abs(d) < 0.8 or not UI.get("motion", True):
                    self.cv.yview_moveto(self.target / total)
                    self.running = False
                    self.target = None
                    return
                self.cv.yview_moveto((cur + d * 0.30) / total)
                self.cv.after(12, tick)
            except tk.TclError:
                self.running = False
        tick()


# ── feedback ─────────────────────────────────────────────────────────────────
class Toast:
    """Non-blocking notification, bottom-right, stacked and self-dismissing."""
    KINDS = {"ok":   (BH_GREEN, ""),
             "info": (BH_TOOL,  ""),
             "warn": (BH_WARN,  ""),
             "err":  (BH_RED,   "")}
    _live = []

    @classmethod
    def show(cls, root, text, kind="info", ms=3200):
        try:
            if not root.winfo_exists():
                return
        except Exception:
            return
        col, icon = cls.KINDS.get(kind, cls.KINDS["info"])
        text = " ".join(str(text).split())
        for old in list(cls._live):         # the newest message replaces the last
            try:
                old.destroy()
            except Exception:
                pass
        cls._live.clear()
        if len(text) > 120:
            text = text[:117] + "…"
        f = tkfont.Font(family=UI["fam"], size=max(8, int(9 * UI["scale"])))
        w = min(max(f.measure(text) + 62, 240), 520)
        h = max(f.metrics("linespace") + 32, 52)
        try:
            tw = tk.Toplevel(root)
            tw.wm_overrideredirect(True)
            tw.configure(bg=BH_INK)
            try:
                tw.wm_attributes("-topmost", True)
                tw.wm_attributes("-alpha", 0.0)
            except tk.TclError:
                pass
            img = _round_img(w, h, R_SM, fill=BH_INK, bg=BH_INK,
                             border=_mix(BH_INK, col, 0.45),
                             key=("toast", w, h, col))
            body = tk.Label(tw, image=img, bd=0, highlightthickness=0, bg=BH_INK)
            body.image = img
            body.pack()
            tk.Frame(tw, bg=col, width=4, height=h - 16).place(x=10, y=8)
            tk.Label(tw, text=text, font=_fnt(9), fg=BH_WHITE, bg=BH_INK,
                     justify=tk.LEFT, anchor="w", wraplength=w - 60).place(
                x=28, y=h // 2, anchor="w")
            root.update_idletasks()
            slot = len(cls._live)
            x0 = root.winfo_rootx() + root.winfo_width() - w - 26
            y0 = root.winfo_rooty() + root.winfo_height() - h - 86 - slot * (h + 10)
            tw.wm_geometry(f"{w}x{h}+{x0 + 26}+{y0}")
            cls._live.append(tw)

            def slide(t):
                tw.wm_geometry(f"{w}x{h}+{int(x0 + 26 * (1 - t))}+{y0}")
                try:
                    tw.wm_attributes("-alpha", t)
                except tk.TclError:
                    pass
            Motion.run(tw, 180, slide, key=("toast-in", id(tw)))

            def close():
                try:
                    if tw in cls._live:
                        cls._live.remove(tw)
                    def out(t):
                        try:
                            tw.wm_attributes("-alpha", 1 - t)
                        except tk.TclError:
                            pass
                    Motion.run(tw, 150, out, key=("toast-out", id(tw)),
                               done=tw.destroy)
                except Exception:
                    pass
            tw.after(ms, close)
            body.bind("<Button-1>", lambda e: close())
        except Exception:
            pass


class _StatusVar(tk.StringVar):
    """Status line that also surfaces as a toast, so progress is visible even
    when the eye is at the top of a long page."""

    def __init__(self, root, value=""):
        super().__init__(value=value)
        self._root = root

    def set(self, value, kind=None, toast=True):
        super().set(value)
        low = str(value).lower()
        if not toast or not value or low in ("ready", "—"):
            return
        if kind is None:
            kind = ("err" if any(s in low for s in ("fail", "error", "could not"))
                    else "warn" if any(s in low for s in ("missing", "not ", "no "))
                    else "ok")
        Toast.show(self._root, value, kind)


class Tooltip:
    """Dark rounded tip that fades in, flips when it would leave the screen."""

    def __init__(self, widget, text, delay=420):
        self.w, self.text, self.delay = widget, text, delay
        self._id = self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _=None):
        self._cancel()
        self._id = self.w.after(self.delay, self._show)

    def _cancel(self):
        if self._id:
            try:
                self.w.after_cancel(self._id)
            except tk.TclError:
                pass
            self._id = None

    def _show(self):
        if self._tip or not self.text:
            return
        try:
            f = tkfont.Font(family=UI["fam"], size=max(7, int(8 * UI["scale"])))
            lines, cur = [], ""
            for word in str(self.text).split():
                probe = (cur + " " + word).strip()
                if f.measure(probe) > 340 and cur:
                    lines.append(cur)
                    cur = word
                else:
                    cur = probe
            if cur:
                lines.append(cur)
            body = "\n".join(lines)
            tw_w = min(max(f.measure(l) for l in lines) + 22, 380)
            tw_h = len(lines) * f.metrics("linespace") + 16

            x = self.w.winfo_rootx() + 12
            y = self.w.winfo_rooty() + self.w.winfo_height() + 8
            sw, sh = self.w.winfo_screenwidth(), self.w.winfo_screenheight()
            if x + tw_w > sw - 12:
                x = max(8, sw - tw_w - 12)
            if y + tw_h > sh - 12:
                y = self.w.winfo_rooty() - tw_h - 8

            self._tip = tw = tk.Toplevel(self.w)
            tw.wm_overrideredirect(True)
            tw.configure(bg=BH_INK)
            try:
                tw.wm_attributes("-topmost", True)
                tw.wm_attributes("-alpha", 0.0)
            except tk.TclError:
                pass
            img = _round_img(tw_w, tw_h, R_SM, fill=BH_INK, bg=BH_INK,
                             border=BH_NAVY_HI, key=("tip", tw_w, tw_h))
            lb = tk.Label(tw, image=img, text=body, compound=tk.CENTER,
                          font=f, fg="#E9EFED", bg=BH_INK, justify=tk.LEFT,
                          bd=0, highlightthickness=0)
            lb.image = img
            lb.pack()
            tw.wm_geometry(f"{tw_w}x{tw_h}+{x}+{y + 4}")

            def step(t):
                try:
                    tw.wm_attributes("-alpha", 0.97 * t)
                except tk.TclError:
                    pass
                tw.wm_geometry(f"{tw_w}x{tw_h}+{x}+{int(y + 4 * (1 - t))}")
            Motion.run(tw, 140, step, key=("tip", id(self)))
        except Exception:
            self._tip = None

    def _hide(self, _=None):
        self._cancel()
        if self._tip:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None


# ─────────────────────────────────────────────────────────────────────────────
# APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
class BudgetOfferApp:

    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} · {APP_SUB} — Baker Hughes Thermodyn · DTI")
        self.root.geometry("1700x960")
        self.root.minsize(1400, 800)
        self.root.configure(bg=BH_DARK)
        self._style()
        self._vars()
        self._load_settings()
        self._build_ui()
        self._on_category()
        self._on_plc()
        self._update_nda()
        self._contacts_load(quiet=True)
        self._contacts_defaults()
        self.root.after(60, lambda: self._scan_bg(silent=True))
        self._resume_offer()
        self._tick_id = self.root.after(800, self._ux_tick)

    def _style(self):
        s = ttk.Style(self.root)
        s.theme_use("clam")

        # Thumb-only scrollbar: no arrows, no trough border, brand on hover.
        try:
            s.layout("BH.Vertical.TScrollbar", [
                ("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
                    ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})
                ]})])
        except tk.TclError:
            pass
        s.configure("BH.Vertical.TScrollbar", troughcolor=BH_BG, borderwidth=0,
                    relief=tk.FLAT, background=BH_GRID, darkcolor=BH_GRID,
                    lightcolor=BH_GRID, bordercolor=BH_BG, width=9)
        s.map("BH.Vertical.TScrollbar",
              background=[("pressed", BH_GREEN), ("active", BH_MUTED)])

        s.configure("TScrollbar", background=BH_GRID, troughcolor=BH_BG,
                    arrowcolor=BH_GRAY, relief=tk.FLAT, borderwidth=0, width=11)
        s.map("TScrollbar", background=[("active", BH_GREEN)])

        s.configure("TNotebook", background=BH_DARK, borderwidth=0,
                    tabmargins=[2, 6, 2, 0])
        s.configure("TNotebook.Tab", font=FONT_B, padding=[18, 8],
                    background=BH_NAVY_HI, foreground=BH_ON_DARK2, borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", BH_GREEN), ("active", BH_LINE_D)],
              foreground=[("selected", BH_WHITE), ("active", BH_WHITE)],
              expand=[("selected", [0, 0, 0, 0])])

        s.configure("BH.TCombobox", arrowcolor=BH_GREEN, borderwidth=1,
                    relief=tk.FLAT, padding=4,
                    fieldbackground=BH_WHITE, background=BH_WHITE,
                    bordercolor=BH_BORDER, lightcolor=BH_BORDER,
                    darkcolor=BH_BORDER, foreground=BH_INK)
        s.map("BH.TCombobox",
              fieldbackground=[("readonly", BH_WHITE), ("focus", BH_WHITE)],
              bordercolor=[("focus", BH_GREEN), ("hover", BH_GREEN)],
              arrowcolor=[("hover", BH_GREEN_D)],
              selectbackground=[("readonly", BH_WHITE)],
              selectforeground=[("readonly", BH_INK)])
        self.root.option_add("*TCombobox*Listbox.font", FONT_N)
        self.root.option_add("*TCombobox*Listbox.background", BH_WHITE)
        self.root.option_add("*TCombobox*Listbox.selectBackground", BH_GREEN)
        self.root.option_add("*TCombobox*Listbox.selectForeground", BH_WHITE)
        self.root.option_add("*TCombobox*Listbox.foreground", BH_INK)

        s.configure("BH.Horizontal.TProgressbar", troughcolor=BH_LINE_D,
                    background=BH_G1, borderwidth=0, thickness=6,
                    lightcolor=BH_G1, darkcolor=BH_G1)
        s.configure("BHLight.Horizontal.TProgressbar", troughcolor=BH_GRID,
                    background=BH_GREEN, borderwidth=0, thickness=6,
                    lightcolor=BH_GREEN, darkcolor=BH_GREEN)

    def _vars(self):
        self.offer_category    = tk.StringVar(value="DTI")
        self.opp_id            = tk.StringVar()
        self.revision          = tk.StringVar()
        self.gen_comment       = tk.StringVar()
        self.application       = tk.StringVar(value=APPS[0])
        self.gas_type          = tk.StringVar(value="Natural Gas")
        self.custom_gas        = tk.StringVar()
        self.client_name       = tk.StringVar()
        self.client_loc        = tk.StringVar()
        self.project_name      = tk.StringVar()
        self.nb_units          = tk.StringVar(value="1")
        self.project_date      = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.ae_name           = tk.StringVar()
        self.ae_phone          = tk.StringVar()
        self.ae_email          = tk.StringVar()
        self.sales_name        = tk.StringVar()
        self.sales_phone       = tk.StringVar()
        self.sales_email       = tk.StringVar()
        self.sales_title       = tk.StringVar()
        self.adv_title         = tk.StringVar()
        self.ae_title          = tk.StringVar()
        self.contacts          = {}      # role -> people read from the Excel list
        self.contacts_file     = tk.StringVar()
        self._w_by_var         = {}      # variable name -> its input widget
        self._file             = ""      # project file being edited
        self._saved_key        = None    # content at the last save or open
        self._autosave_key     = None
        self._autosave_armed   = False
        self._show_off         = False   # lines not in this offer, hidden
        self._ready_sig        = None
        self._tick_id          = None
        self.adv_name          = tk.StringVar()
        self.adv_phone         = tk.StringVar()
        self.adv_email         = tk.StringVar()
        self.nda_signed        = tk.BooleanVar(value=False)
        self.doc_vars          = {}
        # documents to provide: one tick per library document and offer category
        self.lib_vars          = {d["id"]: tk.BooleanVar(value=d["on"]["DTI"])
                                  for d in DB_LIBRARY}
        self.lib_state         = {}      # category -> {document id: ticked}
        self.lib_choice        = {}      # document id -> file picked by hand
        self._lib_cat          = "DTI"
        self._lib_loading      = False
        self._lib_res          = {}
        self._lib_folders      = {}
        for _id, _v in self.lib_vars.items():
            _v.trace_add("write", lambda *_a, i=_id: self._lib_toggled(i))
        self.extra_docs        = []
        self.num_impellers     = tk.StringVar(value="7")
        self.casing_size       = tk.StringVar(value="450")
        self.num_phases        = tk.StringVar(value="2")
        self.config_code       = tk.StringVar()
        self.code_manual       = tk.BooleanVar(value=False)
        self.motor_type        = tk.StringVar(value="MGV1")
        self.back_to_back      = tk.StringVar(value="Back to Back")
        self.rating            = tk.StringVar(value="300")
        self.ucp_type          = tk.StringVar(value="Allen Bradley Common PLC")
        self.plc_option        = tk.StringVar(value="Base Case")
        self.plc_safety        = tk.BooleanVar(value=True)
        self.thrust_bearing    = tk.BooleanVar(value=False)
        self.motor_cooling     = tk.StringVar(value="TCV")     # TCV | Fan
        self.motor_cooling_tcv = tk.BooleanVar(value=True)
        self.motor_cooling_fan = tk.BooleanVar(value=False)
        self.motor_cooling.trace_add("write", lambda *_: (
            self.motor_cooling_tcv.set(self.motor_cooling.get() != "Fan"),
            self.motor_cooling_fan.set(self.motor_cooling.get() == "Fan")))
        self.vfd_cooling       = tk.StringVar(value="Cooler")
        self.transfo_type      = tk.StringVar(value="Oil ONAN")
        self.power             = tk.StringVar()
        self.cooling_system    = tk.StringVar(value="multistage_with_TCV")
        self.antisurge_valve   = tk.BooleanVar(value=False)
        self.filter_type       = tk.StringVar(value="Simplex")
        self.ehouse            = tk.BooleanVar(value=False)
        self.mcc               = tk.BooleanVar(value=False)
        self.comp_enclosure    = tk.BooleanVar(value=False)
        self.cost_base         = tk.StringVar()
        self.margin_pct        = tk.StringVar(value="20")      # CM, % of price
        self.warr_pct          = tk.StringVar(value="0.657")   # warranty, % of price
        self.lds_pct           = tk.StringVar(value="0.64")    # LDs, % of price
        self.final_price       = tk.StringVar()
        self.price_mode        = tk.StringVar(value="auto")    # auto | manual
        self._price_busy       = False
        self._price_note       = tk.StringVar(value="No price yet.")
        self.incoterm          = tk.StringVar(value="FCA")
        self.lead_time         = tk.StringVar()

        # cooling concept, motor cooling and phases stay consistent
        self._sync_guard = False
        self.cooling_system.trace_add("write", self._concept_changed)

        # the configuration code follows the three fields it is made of
        for v in (self.num_impellers, self.casing_size, self.num_phases):
            v.trace_add("write", self._auto_code)
        self._auto_code()

        # Scope rows — one per SCOPE_TABLE entry
        self.scope_rows = [self._scope_row(item) for item in SCOPE_TABLE]
        self.extra_scope_rows = []
        self._scope_only_open = False
        self._scope_sync()
        for v in (self.config_code, self.motor_type, self.transfo_type,
                  self.vfd_cooling, self.thrust_bearing, self.filter_type,
                  self.motor_cooling, self.cooling_system, self.ucp_type,
                  self.plc_option, self.num_phases, self.back_to_back,
                  self.incoterm, self.mcc, self.ehouse, self.antisurge_valve,
                  self.comp_enclosure):
            v.trace_add("write", self._scope_sync)

        # ── document manager ──────────────────────────────────────────────
        base = _tool_dir()
        self.user_root   = tk.StringVar()
        self.db_root     = tk.StringVar(value=find_library(base)
                                    or os.path.join(base, "DB-2026"))
        self._project_pick = ""     # folder picked, maybe above the proposal
        self.use_db      = tk.BooleanVar(value=True)   # global DB switch
        self.auto_place  = tk.BooleanVar(value=True)   # copy DB docs into project
        self.doc_use_db  = {}       # code -> BooleanVar (per document)
        self.doc_status  = {}       # code -> dict(state, path, …)
        self.doc_attach  = {}       # code -> [files picked by hand]
        self.doc_layout  = {}       # "FIRM" / "DTI" -> section and item order
        self._struct_cache = (None, None)
        self.pdf_output  = tk.StringVar(value="single")   # single | folder
        self.pdf_use_word = tk.BooleanVar(value=True)
        self.pdf_frame   = tk.BooleanVar(value=True)
        self.skeleton_file = tk.StringVar()
        self.pdf_placeholders = tk.BooleanVar(value=True)
        self.pdf_stamp        = tk.BooleanVar(value=True)
        self.pdf_scope        = tk.BooleanVar(value=True)
        self.pdf_commercial   = tk.BooleanVar(value=True)
        self.pdf_skip_empty   = tk.BooleanVar(value=False)
        self.doc_filter  = tk.StringVar(value="all")
        self.doc_search  = tk.StringVar()
        self._doc_collapsed = set()
        # configuration changes alter which standard document applies
        for v in (self.cooling_system, self.casing_size, self.ucp_type,
                  self.plc_option, self.motor_type, self.config_code):
            v.trace_add("write", self._refresh_docs)

    # ── one scrollable page + its entry in the navigation rail ────────────────
    def _page(self, num, title, hint):
        outer  = tk.Frame(self._pages, bg=BH_BG)
        canvas = tk.Canvas(outer, bg=BH_BG, highlightthickness=0, bd=0)
        vsb    = ttk.Scrollbar(outer, orient="vertical",
                               style="BH.Vertical.TScrollbar",
                               command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 3), pady=6)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner  = tk.Frame(canvas, bg=BH_BG)
        wid    = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))

        scroller = _SmoothScroll(canvas)

        def _wheel(e):
            scroller.by(-(e.delta / 120.0) * 112)

        def _wheel_x11(e):
            scroller.by(-112 if e.num == 4 else 112)

        def _bind(_=None):
            canvas.bind_all("<MouseWheel>", _wheel)
            canvas.bind_all("<Button-4>", _wheel_x11)
            canvas.bind_all("<Button-5>", _wheel_x11)

        def _unbind(_=None):
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                canvas.unbind_all(seq)
        outer.bind("<Enter>", _bind)
        outer.bind("<Leave>", _unbind)
        inner._canvas = canvas
        inner._scroller = scroller
        inner._win = wid

        # ── page banner: step number, title, intent ──────────────────────────
        head = tk.Frame(inner, bg=BH_BG)
        head.pack(fill=tk.X, padx=SP4, pady=(SP4, SP3))
        tx = tk.Frame(head, bg=BH_BG)
        tx.pack(side=tk.LEFT)
        _lbl(tx, f"{num}   {title.upper()}", fnt=FONT_S, fg=BH_INK,
             bg=BH_BG).pack(anchor=tk.W)
        _lbl(tx, hint, fnt=FONT_SM, fg=BH_MUTED, bg=BH_BG).pack(anchor=tk.W,
                                                                pady=(3, 0))
        step = _lbl(head, "", fnt=FONT_XS, fg=BH_MUTED, bg=BH_BG)
        step.pack(side=tk.RIGHT, pady=(4, 0))
        tk.Frame(inner, bg=BH_BORDER, height=1).pack(fill=tk.X, padx=SP4,
                                                     pady=(0, SP3))

        # ── rail entry ───────────────────────────────────────────────────────
        idx  = len(self._pages_list)
        item = tk.Frame(self._rail_items, bg=BH_INK, cursor="hand2")
        item.pack(fill=tk.X)
        accent = tk.Frame(item, bg=BH_INK, width=3)
        accent.pack(side=tk.LEFT, fill=tk.Y)
        txt = tk.Frame(item, bg=BH_INK, padx=15, pady=10)
        txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
        n = tk.Label(txt, text=num,   font=FONT_CS, bg=BH_INK, fg=BH_MUTED, anchor=tk.W)
        t = tk.Label(txt, text=title, font=FONT_B,  bg=BH_INK, fg=BH_ON_DARK2, anchor=tk.W)
        h = tk.Label(txt, text=hint,  font=FONT_XS, bg=BH_INK, fg=BH_MUTED, anchor=tk.W)
        for w in (n, t, h):
            w.pack(anchor=tk.W)
        badge = tk.Label(item, text="", font=FONT_BD, bg=BH_INK, fg=BH_MUTED,
                         padx=14)
        badge.pack(side=tk.RIGHT)
        widgets = (item, txt, n, t, h, badge)
        for w in widgets:
            w.bind("<Button-1>", lambda e, i=idx: self._show_page(i))
        page = dict(outer=outer, accent=accent, widgets=widgets, num=n, title=t,
                    item=item, inner=inner, step=step, idx=idx, hovered=False,
                    badge=badge, name=title)
        item.bind("<Enter>", lambda e, pg=page: self._rail_hover(pg, True), add="+")
        item.bind("<Leave>", lambda e, pg=page: self._rail_hover(pg, False), add="+")
        Tooltip(item, f"{title} — {hint}    (Ctrl+{idx + 1})")
        self._pages_list.append(page)
        return inner

    def _rail_hover(self, page, on):
        if page["idx"] == getattr(self, "_page_idx", 0):
            return
        ws = [w for w in page["widgets"] if w is not page["accent"]]
        _fade(ws, BH_INK if on else BH_NAVY_HI, BH_NAVY_HI if on else BH_INK,
              120, key=("rail", page["idx"]))
        _fade_fg(page["title"], BH_ON_DARK2 if on else BH_WHITE,
                 BH_WHITE if on else BH_ON_DARK2, 120)

    def _show_page(self, i):
        i = max(0, min(i, len(self._pages_list) - 1))
        if getattr(self, "_page_idx", None) == i and \
           self._pages_list[i]["outer"].winfo_ismapped():
            return
        for j, pg in enumerate(self._pages_list):
            on = (i == j)
            Motion.stop(("rail", j))
            pg["outer"].pack_forget()
            bg = BH_NAVY_HI if on else BH_INK
            for w in pg["widgets"]:
                if w is not pg["accent"]:
                    w.config(bg=bg)
            pg["title"].config(fg=BH_WHITE if on else BH_ON_DARK2)
            pg["num"].config(fg=BH_ACCENT if on else BH_MUTED)
        page = self._pages_list[i]
        page["outer"].pack(fill=tk.BOTH, expand=True)
        self._page_idx = i

        self._rail_sync(animate=True)

        # content eases up into place
        inner = page["inner"]
        cv, wid = inner._canvas, inner._win
        cv.yview_moveto(0)
        Motion.run(cv, 190, lambda t: cv.coords(wid, 0, int(16 * (1 - t))),
                   key=("page", id(cv)),
                   done=lambda: cv.configure(scrollregion=cv.bbox("all")))

    def _focus_search(self):
        """Ctrl+F: jump to the document search, wherever the user is."""
        se = getattr(self, "_doc_search_entry", None)
        if se is None:
            return
        for i, pg in enumerate(self._pages_list):
            if str(se).startswith(str(pg["outer"])):
                self._show_page(i)
                break
        self.root.after_idle(lambda: (se.focus_set(), se.select_range(0, tk.END)))

    def _escape(self):
        """Esc: close the open hint, clear the search, nothing destructive."""
        if getattr(self, "doc_search", None) is not None and self.doc_search.get():
            self.doc_search.set("")
            try:
                self._apply_visibility()
            except Exception:
                pass

    def _rail_sync(self, animate=False):
        """Keep the sliding rail indicator glued to the active step, including
        the first paint, when the window is not mapped yet and heights are 1."""
        ind = getattr(self, "_rail_ind", None)
        pages = getattr(self, "_pages_list", None)
        if ind is None or not pages:
            return
        item = pages[max(0, min(getattr(self, "_page_idx", 0), len(pages) - 1))]["item"]
        try:
            self._rail_items.update_idletasks()
            h = item.winfo_height()
            if h <= 1:
                h = item.winfo_reqheight()
            y = item.winfo_y()
            y0 = ind.winfo_y() if (animate and ind.winfo_ismapped()) else y
            ind.place(x=0, y=y0, width=4, height=h)
            ind.lift()
            if animate and y0 != y:
                Motion.run(ind, 220,
                           lambda t: ind.place_configure(y=int(y0 + (y - y0) * t),
                                                         height=h),
                           key="railind", ease=Motion.ease_in_out)
        except tk.TclError:
            pass

    # ── section card ──────────────────────────────────────────────────────────
    def _section(self, parent, title, sub=""):
        """Plain card: hairline outline, one thin brand rule on the header.
        Returns the body frame; `body.shell` hides the whole card."""
        shell = tk.Frame(parent, bg=BH_BG)
        shell.pack(fill=tk.X, padx=SP4, pady=(0, SP3))

        frame = tk.Frame(shell, bg=BH_CARD_BR)
        frame.pack(fill=tk.X)
        card = tk.Frame(frame, bg=BH_WHITE)
        card.pack(fill=tk.X, padx=1, pady=1)
        tk.Frame(shell, bg=BH_SHADOW_2, height=1).pack(fill=tk.X)

        head = tk.Frame(card, bg=BH_LIGHT)
        head.pack(fill=tk.X)
        tk.Frame(head, bg=BH_GREEN, width=3).pack(side=tk.LEFT, fill=tk.Y)
        _lbl(head, f"   {title.upper()}", fnt=FONT_BD, fg=BH_INK, bg=BH_LIGHT,
             pady=10).pack(side=tk.LEFT)
        if sub:
            _lbl(head, f"{sub}   ", fnt=FONT_SM, fg=BH_MUTED,
                 bg=BH_LIGHT).pack(side=tk.RIGHT)
        tk.Frame(card, bg=BH_BORDER, height=1).pack(fill=tk.X)

        body = tk.Frame(card, bg=BH_WHITE, padx=SP4, pady=SP3)
        body.pack(fill=tk.X)
        body.shell = shell          # so a whole card can be hidden/shown
        return body

    # ─────────────────────────────────────────────────────────────────────────
    # BUILD UI
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._header()
        body = tk.Frame(self.root, bg=BH_BG)
        body.pack(fill=tk.BOTH, expand=True)

        rail = tk.Frame(body, bg=BH_INK, width=248)
        rail.pack(side=tk.LEFT, fill=tk.Y)
        rail.pack_propagate(False)

        cap = tk.Frame(rail, bg=BH_INK)
        cap.pack(fill=tk.X, padx=18, pady=(14, 6))
        _lbl(cap, "WORKFLOW", fnt=FONT_XS, fg=BH_MUTED,
             bg=BH_INK).pack(anchor=tk.W)

        self._rail_items = tk.Frame(rail, bg=BH_INK)
        self._rail_items.pack(fill=tk.X, pady=(4, 0))
        self._rail_ind = tk.Frame(self._rail_items, bg=BH_ACCENT, width=4)
        self._rail_items.bind("<Configure>", lambda e: self._rail_sync())

        # readiness panel at the foot of the rail
        foot = tk.Frame(rail, bg=BH_NAVY_2, padx=16, pady=14)
        foot.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Frame(foot, bg=BH_LINE_D, height=1).pack(fill=tk.X, pady=(0, 12))
        rt = tk.Frame(foot, bg=BH_NAVY_2)
        rt.pack(fill=tk.X)
        _lbl(rt, "Document package", fnt=FONT_XS, fg=BH_MUTED,
             bg=BH_NAVY_2).pack(anchor=tk.W)
        self._rail_ready = tk.StringVar(value="—")
        tk.Label(rt, textvariable=self._rail_ready, font=FONT_S,
                 fg=BH_WHITE, bg=BH_NAVY_2).pack(anchor=tk.W, pady=(2, 6))
        self._rail_bar = ttk.Progressbar(rt, style="BH.Horizontal.TProgressbar",
                                         mode="determinate", maximum=100,
                                         length=180)
        self._rail_bar.pack(fill=tk.X)

        hint = _mix(BH_MUTED, BH_NAVY_2, 0.30)
        _lbl(foot, "Ctrl+1…6   pages", fnt=FONT_XS, fg=hint,
             bg=BH_NAVY_2).pack(anchor=tk.W, pady=(12, 0))
        _lbl(foot, "Ctrl+G  PDF      Ctrl+S  save      F5  scan",
             fnt=FONT_XS, fg=hint, bg=BH_NAVY_2).pack(anchor=tk.W)

        self._pages = tk.Frame(body, bg=BH_BG)
        self._pages.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._pages_list = []

        self.t_general    = self._page("01", "General information",   "Client, contacts, NDA")
        self.t_technical  = self._page("02", "Technical configuration","Compressor, driver, cooling")
        self.t_docs       = self._page("03", "Documents to provide",   "Offer package checklist")
        self.t_scope      = self._page("04", "Scope of supply",        "BH / customer split")
        self.t_commercial = self._page("05", "Commercial conditions",  "Price, INCOTERM, lead time")
        self.t_docmgr     = self._page("06", "Document manager",       "Assemble and generate the PDF")

        total = len(self._pages_list)
        for k, pg in enumerate(self._pages_list, 1):
            pg["step"].config(text=f"STEP {k} OF {total}")

        self._tab_general()
        self._tab_technical()
        self._tab_docs()
        self._tab_scope()
        self._tab_commercial()
        self._tab_docmgr()
        self._action_bar()
        self._shortcuts()
        self._page_nav()
        self._show_page(0)

    def _shortcuts(self):
        """One keyboard path to everything that matters."""
        r = self.root
        for k in range(1, 7):
            r.bind(f"<Control-Key-{k}>", lambda e, i=k - 1: self._show_page(i))
        r.bind("<Control-g>", lambda e: self._generate_pdf())
        r.bind("<Control-e>", lambda e: self._export_excel())
        r.bind("<Control-s>", lambda e: self._save_json())
        r.bind("<Control-S>", lambda e: self._save_json(ask=True))
        r.bind("<Control-o>", lambda e: self._load_json())
        r.bind("<Control-n>", lambda e: self._reset())
        r.protocol("WM_DELETE_WINDOW", self._on_close)
        # Enter moves to the next field, as in a form
        r.bind_class("Entry", "<Return>",
                     lambda e: (e.widget.tk_focusNext().focus_set(), "break")[1])
        r.bind("<F5>",        lambda e: self._scan_bg(silent=False))
        r.bind("<Control-f>", lambda e: self._focus_search())
        r.bind("<Escape>",    lambda e: self._escape(), add="+")
        r.bind("<Control-Next>",  lambda e: self._show_page(
            getattr(self, "_page_idx", 0) + 1))
        r.bind("<Control-Prior>", lambda e: self._show_page(
            getattr(self, "_page_idx", 0) - 1))

    # ── HEADER ────────────────────────────────────────────────────────────────
    def _header(self):
        wrap = tk.Frame(self.root, bg=BH_DARK)
        wrap.pack(fill=tk.X)
        hdr = tk.Frame(wrap, bg=BH_DARK, height=86)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)

        # wordmark — no badge, no monogram
        lf = tk.Frame(hdr, bg=BH_DARK)
        lf.pack(side=tk.LEFT, padx=(26, 0), pady=16)
        tk.Frame(lf, bg=BH_GREEN, width=3).pack(side=tk.LEFT, fill=tk.Y,
                                                pady=2)
        tf = tk.Frame(lf, bg=BH_DARK)
        tf.pack(side=tk.LEFT, padx=(14, 0))
        _lbl(tf, APP_NAME, fnt=_fnt(21, True), fg=BH_WHITE,
             bg=BH_DARK).pack(anchor=tk.W)
        _lbl(tf, f"{APP_SUB}   ·   {APP_OWNER}", fnt=FONT_SM, fg=BH_ON_DARK,
             bg=BH_DARK).pack(anchor=tk.W, pady=(3, 0))

        # Baker Hughes signature
        bh = tk.Frame(hdr, bg=BH_DARK)
        bh.pack(side=tk.LEFT, padx=26, pady=22)
        tk.Frame(bh, bg=BH_LINE_D, width=1).pack(side=tk.LEFT, fill=tk.Y,
                                                 padx=(0, 22))
        try:
            lp = os.path.join(_tool_dir(), "baker_hughes_logo.png")
            if os.path.exists(lp):
                img = Image.open(lp)
                ratio = img.height / max(img.width, 1)
                w = int(164 * UI["scale"])
                img = img.resize((w, max(int(w * ratio), 1)),
                                 Image.Resampling.LANCZOS)
                self._logo = ImageTk.PhotoImage(img)
                tk.Label(bh, image=self._logo, bg=BH_DARK).pack(side=tk.LEFT)
            else:
                raise FileNotFoundError
        except Exception:
            _lbl(bh, "Baker Hughes", fnt=_fnt(13, True), fg=BH_ON_DARK2,
                 bg=BH_DARK).pack(side=tk.LEFT)

        # right side: CCS import + offer category
        pf = tk.Frame(hdr, bg=BH_DARK)
        pf.pack(side=tk.RIGHT, padx=26, pady=16)
        cat = tk.Frame(pf, bg=BH_DARK)
        cat.pack(side=tk.RIGHT, padx=(26, 0))
        _lbl(cat, "OFFER CATEGORY", fnt=FONT_XS, fg=BH_MUTED,
             bg=BH_DARK).pack(anchor=tk.E)

        seg = tk.Frame(cat, bg=BH_LINE_D)
        seg.pack(anchor=tk.E, pady=(6, 0))
        self._cat_btns = {}
        for val, col in [("DTI", BH_GREEN), ("BUDGET", BH_WARN),
                         ("FIRM", BH_RED)]:
            b = tk.Label(seg, text=val, font=FONT_B, bg=BH_NAVY_HI,
                         fg=BH_ON_DARK2, cursor="hand2",
                         padx=16, pady=6)
            b.pack(side=tk.LEFT, padx=1, pady=1)
            b.bind("<Button-1>", lambda e, v=val: self._set_category(v))
            b.bind("<Enter>", lambda e, wd=b, v=val: wd.config(fg=BH_WHITE)
                   if v != self.offer_category.get() else None)
            b.bind("<Leave>", lambda e, wd=b, v=val: wd.config(
                fg=BH_WHITE if v == self.offer_category.get() else BH_ON_DARK2))
            self._cat_btns[val] = (b, col)

        self._ccs_button(pf).pack(side=tk.RIGHT)

        tk.Frame(wrap, bg=BH_GREEN, height=2).pack(fill=tk.X)

        # live context strip
        strip = tk.Frame(wrap, bg=BH_NAVY_2, height=30)
        strip.pack(fill=tk.X)
        strip.pack_propagate(False)
        self._ctx = tk.StringVar()
        tk.Label(strip, textvariable=self._ctx, font=FONT_SM, bg=BH_NAVY_2,
                 fg=BH_ON_DARK, anchor=tk.W).pack(side=tk.LEFT, padx=26)
        tk.Label(strip, text=f"v{APP_VER}", font=FONT_XS, bg=BH_NAVY_2,
                 fg=BH_MUTED).pack(side=tk.RIGHT, padx=26)
        tk.Frame(wrap, bg=BH_LINE_D, height=1).pack(fill=tk.X)
        for v in (self.client_name, self.opp_id, self.revision, self.config_code,
                  self.application, self.motor_type, self.rating, self.incoterm,
                  self.offer_category):
            v.trace_add("write", self._ctx_update)
        self._paint_category()
        self._ctx_update()

    def _ccs_button(self, parent):
        """Plain outlined action: reads the project data out of a CCS output."""
        base, hov = BH_DARK, BH_NAVY_HI
        f = tk.Frame(parent, bg=base, cursor="hand2", padx=14, pady=9,
                     highlightthickness=1, highlightbackground=BH_LINE_D,
                     highlightcolor=BH_LINE_D)
        tx = tk.Frame(f, bg=base)
        tx.pack(side=tk.LEFT)
        l1 = tk.Label(tx, text="Load from CCS", font=FONT_B, bg=base,
                      fg=BH_WHITE, anchor=tk.W)
        l2 = tk.Label(tx, text="datasheet or export file", font=FONT_XS, bg=base,
                      fg=BH_ON_DARK, anchor=tk.W)
        l1.pack(anchor=tk.W)
        l2.pack(anchor=tk.W)
        widgets = [f, tx, l1, l2]
        for w in widgets:
            w.bind("<Button-1>", lambda e: self._load_ccs())
            w.bind("<Enter>", lambda e: [x.config(bg=hov) for x in widgets])
            w.bind("<Leave>", lambda e: [x.config(bg=base) for x in widgets])
        Tooltip(f, "Read a CCS datasheet or export file and fill in the customer, "
                   "the job number, the machine model and the duty point.")
        return f

    # ─────────────────────────────────────────────────────────────────────────
    # CCS IMPORT
    # ─────────────────────────────────────────────────────────────────────────
    def _load_ccs(self):
        path = filedialog.askopenfilename(
            title="Select a CCS datasheet or export file",
            initialdir=self.user_root.get().strip() or None,
            filetypes=[("CCS files", "*.pdf *.txt *.out *.csv *.json *.dat"),
                       ("All files", "*.*")])
        if not path:
            return
        try:
            text = self._read_ccs_text(path)
        except Exception as e:
            messagebox.showerror("CCS import", f"The file could not be read.\n\n{e}")
            return
        if not text.strip():
            messagebox.showwarning(
                "CCS import",
                "No readable text was found in this file.\n\n"
                "A scanned datasheet has no text layer — export the CCS data as "
                "a text file, or use the searchable PDF.")
            return
        found = self._parse_ccs(text)
        if not found:
            messagebox.showinfo(
                "CCS import",
                "The file was read but none of the expected CCS labels were "
                "found (CUSTOMER, JOB, MODEL, SERVICE…).\n\n"
                "Check that this is a CCS output for an ICL machine.")
            return
        self._ccs_dialog(found, path)

    def _read_ccs_text(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            if not HAS_PYPDF:
                raise RuntimeError("Reading a PDF datasheet needs pypdf "
                                   "(pip install pypdf).")
            pages = PdfReader(path).pages
            texts = []
            for pg in pages[:80]:
                try:
                    texts.append(pg.extract_text() or "")
                except Exception:
                    texts.append("")
            # the datasheet itself carries the labels we look for
            ds = [t for t in texts
                  if "DATASHEET" in t.upper() or "PLANT LOCATION" in t.upper()]
            return "\n".join((ds[:8] or texts[:10]) + texts[:2])
        if ext == ".json":
            with open(path, encoding="utf-8", errors="ignore") as fh:
                data = json.load(fh)
            return "\n".join(f"{k}: {v}" for k, v in _flatten_json(data))
        with open(path, encoding="utf-8", errors="ignore") as fh:
            return fh.read()

    def _parse_ccs(self, text):
        """Pull the fields Y-B-TOOL knows about out of a CCS output."""
        import re
        flat = re.sub(r"[ \t]+", " ", text)
        lines = [l.strip() for l in flat.splitlines() if l.strip()]
        joined = "\n".join(lines)
        found = {}

        STOP = ("NO. REQUIRED", "ITEM NO", "INQUIRY NO", "MANUFACTURER",
                "DRIVER TYPE", "PURCHASE ORDER", "TITLE:", "DOCUMENT CODE",
                "REVISION", "ORIGINAL JOB", "SECURITY CODE", "PAGE ", "SHEET")

        def clean(val):
            up = val.upper()
            cut = len(val)
            for tok in STOP:
                i = up.find(tok)
                if 0 < i < cut:
                    cut = i
            return re.sub(r"\s+", " ", val[:cut]).strip(" :.-|\u2014")

        def grab(patterns, flags=re.I):
            for pat in patterns:
                m = re.search(pat, joined, flags)
                if m:
                    val = clean(m.group(1) or "")
                    if val and val.lower() not in ("n/a", "na", "-", "x"):
                        return val
            return None

        v = grab([r"CUSTOMER\s*NAME\s*[:\n ]\s*(.+)",
                  r"CUSTOMER\s*[:\n ]\s*(.+)", r"^\s*FOR\s+(.+)$"])
        if v: found["client_name"] = (v, self.client_name, "Client name")
        v = grab([r"PLANT LOCATION\s*[:\n ]\s*(.+)", r"^\s*SITE\s+(.+)$"])
        if v: found["client_loc"] = (v, self.client_loc, "Location")
        v = grab([r"\bJOB\b\s*[:\n ]\s*([0-9]{4,10})",
                  r"INQUIRY NO\.?\s*[:\n ]?\s*([0-9]{4,10})",
                  r"ORIGINAL JOB\s*[:\n ]?\s*([0-9]{4,10})"])
        if v: found["opp_id"] = (v, self.opp_id, "OPP / job number")
        v = grab([r"\bSERVICE\b\s*[:\n ]\s*(.+)",
                  r"^\s*PLANT\s*[:\n ]\s*(.+)$"])
        if v: found["service"] = (v, None, "Service (CCS)")
        v = grab([r"ITEM NO\.?\s*[:\n ]?\s*([A-Z0-9]{3,10})\b"])
        if v: found["item_no"] = (v, None, "Item number (CCS)")

        # machine model, e.g. ICL504 / 2ICL358 / 2ICL408B
        m = re.search(r"\b([0-9]?)ICL\s?([0-9]{2,3})([0-9])([A-Za-z]?)\b", joined)
        if m:
            phases, casing, imp = m.group(1) or "1", m.group(2), m.group(3)
            found["model"] = (f"{m.group(1)}ICL{casing}{imp}{m.group(4)}".lstrip(),
                              None, "Machine model")
            found["casing_size"] = (casing + "0", self.casing_size, "Casing size")
            found["num_impellers"] = (imp, self.num_impellers, "Number of impellers")
            found["num_phases"] = (phases, self.num_phases, "Number of phases")

        # duty figures: only when the value sits on the label's own line,
        # otherwise a tabulated datasheet would hand back a neighbouring cell
        v = grab([r"EXP\.? POWER @ MOTOR SHAFT END[^\n]*?\b([0-9]{2,6})\b",
                  r"(?:RATED|ABSORBED|EXPECTED)[^\n]*POWER[^\n]*?\b([0-9]{2,6})\b",
                  r"\bPOWER\s*\(?kW\)?\s*[:=]\s*([0-9]{2,6})"])
        if v and 50 <= _num(v) <= 100000:
            found["power"] = (v, self.power, "Power (kW)")
        v = grab([r"\bSPEED\b[^\n]*?\b([0-9]{3,6})\s*(?:RPM)?\b",
                  r"\bRPM\s*[:=]\s*([0-9]{3,6})"])
        if v and 500 <= _num(v) <= 30000:
            found["speed"] = (v, None, "Speed (rpm)")
        v = grab([r"MOLECULAR WEIGHT\s*[:\n ]?\s*([0-9]+[.,][0-9]+)",
                  r"AVG\.? MOL\.? WT\.?\s*[:\n ]?\s*([0-9]+[.,][0-9]+)"])
        if v: found["mw"] = (v, None, "Molecular weight")
        v = grab([r"GAS HANDLED\s*(?:\(.*?\))?\s*[:\n ]?\s*([A-Za-z][^\n]*)"])
        if v: found["gas"] = (v, None, "Gas handled")
        v = grab([r"\bDATE\b\s*[:\n]?\s*([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4})"])
        if v: found["project_date"] = (v, self.project_date, "Date")
        return found

    def _ccs_dialog(self, found, path):
        win = tk.Toplevel(self.root)
        win.title("CCS import")
        win.configure(bg=BH_WHITE)
        win.geometry("720x560")
        win.transient(self.root)
        win.grab_set()

        hd = tk.Frame(win, bg=BH_DARK, padx=20, pady=14)
        hd.pack(fill=tk.X)
        row = tk.Frame(hd, bg=BH_DARK); row.pack(fill=tk.X)
        ic = tk.Canvas(row, width=34, height=34, bg=BH_DARK,
                       highlightthickness=0, bd=0)
        _round_grad(ic, 1, 1, 33, 33, 8, BH_G1, BH_G3)
        ic.create_text(17, 17, text="CCS", font=("Segoe UI", 9, "bold"),
                       fill=BH_WHITE)
        ic.pack(side=tk.LEFT, padx=(0, 12))
        tx = tk.Frame(row, bg=BH_DARK); tx.pack(side=tk.LEFT)
        _lbl(tx, "Data read from CCS", fnt=("Segoe UI", 13, "bold"),
             fg=BH_WHITE, bg=BH_DARK).pack(anchor=tk.W)
        _lbl(tx, os.path.basename(path), fnt=FONT_SM, fg=BH_ON_DARK,
             bg=BH_DARK).pack(anchor=tk.W)
        _grad_bar(win, 3, bg=BH_WHITE).pack(fill=tk.X)

        body = tk.Frame(win, bg=BH_WHITE, padx=18, pady=14)
        body.pack(fill=tk.BOTH, expand=True)
        _lbl(body, "Tick what you want to bring into the offer. Nothing is "
                   "overwritten until you apply.",
             fnt=FONT_SM, fg=BH_GRAY).pack(anchor=tk.W, pady=(0, 10))

        head = tk.Frame(body, bg=BH_GREEN)
        head.pack(fill=tk.X)
        for txt, w in [("", 3), ("FIELD", 24), ("CURRENT VALUE", 24),
                       ("FROM CCS", 30)]:
            _lbl(head, txt, fnt=FONT_BD, fg=BH_WHITE, bg=BH_GREEN, width=w,
                 anchor=tk.W, pady=5).pack(side=tk.LEFT)

        picks = {}
        for i, (key, (val, var, label)) in enumerate(found.items()):
            bg = BH_WHITE if i % 2 == 0 else BH_LIGHT
            r = tk.Frame(body, bg=bg); r.pack(fill=tk.X)
            applies = var is not None
            picks[key] = tk.BooleanVar(value=applies)
            holder = tk.Frame(r, bg=bg, width=26, height=22)
            holder.pack(side=tk.LEFT); holder.pack_propagate(False)
            if applies:
                tk.Checkbutton(holder, variable=picks[key], bg=bg,
                               selectcolor=BH_WHITE, bd=0, highlightthickness=0,
                               activebackground=bg).pack()
            _lbl(r, label, fnt=FONT_N, bg=bg, width=24, anchor=tk.W).pack(
                side=tk.LEFT, pady=3)
            cur = (var.get() if var is not None else "")
            _lbl(r, (cur or "—")[:30], fnt=FONT_SM,
                 fg=BH_INK if cur else BH_MUTED, bg=bg, width=24,
                 anchor=tk.W).pack(side=tk.LEFT)
            _lbl(r, str(val)[:40], fnt=FONT_B,
                 fg=BH_GREEN if applies else BH_GRAY, bg=bg,
                 anchor=tk.W).pack(side=tk.LEFT, fill=tk.X, expand=True)

        info = [f"{lab}: {val}" for k, (val, var, lab) in found.items()
                if var is None]
        if info:
            box = tk.Frame(body, bg=BH_SEC, padx=12, pady=9,
                           highlightbackground=BH_BORDER, highlightthickness=1)
            box.pack(fill=tk.X, pady=(14, 0))
            _lbl(box, "Read for information, no matching field in the tool:",
                 fnt=FONT_BD, fg=BH_INK, bg=BH_SEC).pack(anchor=tk.W)
            _lbl(box, "     ·     ".join(info), fnt=FONT_SM, fg=BH_GRAY,
                 bg=BH_SEC, wraplength=640, justify=tk.LEFT).pack(anchor=tk.W)

        ft = tk.Frame(win, bg=BH_LIGHT, padx=18, pady=12)
        ft.pack(fill=tk.X)

        def apply_now():
            n = 0
            for key, (val, var, _l) in found.items():
                if var is not None and picks[key].get():
                    var.set(str(val)); n += 1
            self._auto_code()
            self._render_scope()
            win.destroy()
            self._status.set(f"CCS import — {n} field(s) filled in.")
            messagebox.showinfo("CCS import",
                                f"{n} field(s) filled in from {os.path.basename(path)}.")

        _btn(ft, "Apply to the offer", apply_now).pack(side=tk.LEFT)
        _btn(ft, "Cancel", win.destroy, bg=BH_GRAY).pack(side=tk.RIGHT)

    def _set_category(self, val):
        self.offer_category.set(val)
        self._paint_category()
        self._on_category()

    def _paint_category(self):
        """Flat segmented control: the active category is filled, the others
        stay recessed. No sliding pill, no gradient."""
        cur = self.offer_category.get()
        for val, (b, col) in getattr(self, "_cat_btns", {}).items():
            on = (val == cur)
            b.config(bg=col if on else BH_NAVY_HI,
                     fg=_on(col) if on else BH_ON_DARK2, font=FONT_B)

    def _ctx_update(self, *_):
        if not hasattr(self, "_ctx"):
            return
        bits = [self.client_name.get().strip() or "No client yet",
                f"OPP {self.opp_id.get().strip() or '—'}"
                + (f"  rev {self.revision.get().strip()}" if self.revision.get().strip() else ""),
                self.config_code.get().strip() or "no configuration",
                self.application.get(),
                f"{self.motor_type.get()}   ANSI {self.rating.get()}   {self.incoterm.get()}"]
        self._ctx.set("      |      ".join(bits))

    # ── TAB 1: GENERAL ────────────────────────────────────────────────────────
    def _tab_general(self):
        p = self.t_general

        f = self._section(p, "Project Information")
        g = tk.Frame(f, bg=BH_WHITE); g.pack(fill=tk.X)
        self._kv(g, 0, 0, "OPP ID",         self.opp_id,      28)
        self._kv(g, 0, 2, "Revision Number", self.revision,    18)
        self._kv(g, 1, 0, "General Comment", self.gen_comment, 72, cs=3)

        f = self._section(p, "Application & Gas")
        g = tk.Frame(f, bg=BH_WHITE); g.pack(fill=tk.X)
        _lbl(g,"Application Type:",fnt=FONT_B).grid(row=0,column=0,sticky=tk.W,padx=(0,6),pady=4)
        ac=_combo(g,self.application,APPS,w=30)
        ac.grid(row=0,column=1,sticky=tk.W,padx=(0,30),pady=4); ac.current(0)
        _lbl(g,"Gas Type:",fnt=FONT_B).grid(row=0,column=2,sticky=tk.W,padx=(0,6))
        gf=tk.Frame(g,bg=BH_WHITE); gf.grid(row=0,column=3,sticky=tk.W)
        for txt,val in [("Natural Gas","Natural Gas"),("Other:","Custom")]:
            tk.Radiobutton(gf,text=txt,variable=self.gas_type,value=val,font=FONT_N,
                           bg=BH_WHITE,selectcolor=BH_WHITE,activebackground=BH_WHITE,
                           command=self._toggle_gas).pack(side=tk.LEFT,padx=4)
        self._gas_entry=_entry(gf,self.custom_gas,w=18,state=tk.DISABLED)
        self._gas_entry.pack(side=tk.LEFT,padx=4)

        f = self._section(p, "Client Information")
        g = tk.Frame(f, bg=BH_WHITE); g.pack(fill=tk.X)
        self._kv(g,0,0,"Client Name",  self.client_name,32)
        self._kv(g,0,2,"Location",      self.client_loc, 32)
        self._kv(g,1,0,"Date",          self.project_date,18)
        self._kv(g,1,2,"Project name",  self.project_name,32)
        self._kv(g,2,0,"Number of lines", self.nb_units,8)

        f = self._section(p, "Contacts")
        src = tk.Frame(f, bg=BH_WHITE); src.pack(fill=tk.X, pady=(0, 10))
        self._contacts_lbl = _lbl(src, "", fnt=FONT_SM, fg=BH_GRAY)
        self._contacts_lbl.pack(side=tk.LEFT)
        self._contacts_btns = tk.Frame(src, bg=BH_WHITE)
        self._contacts_btns.pack(side=tk.RIGHT)
        cf = tk.Frame(f, bg=BH_WHITE); cf.pack(fill=tk.X)
        self._contact_boxes = {}
        # same order as on the cover of the proposal
        self._contact(cf,0,"Sales Manager",
                      self.sales_name,self.sales_phone,self.sales_email,"sales")
        tk.Frame(cf,bg=BH_BORDER,width=1).grid(row=0,column=1,sticky=tk.NS,padx=18)
        self._contact(cf,2,"Sales Advisor",
                      self.adv_name,self.adv_phone,self.adv_email,"adv")
        tk.Frame(cf,bg=BH_BORDER,width=1).grid(row=0,column=3,sticky=tk.NS,padx=18)
        self._contact(cf,4,"Application Engineer (AE)",
                      self.ae_name,self.ae_phone,self.ae_email,"ae")

        f = self._section(p, "Non-Disclosure Agreement")
        nf=tk.Frame(f,bg=BH_WHITE); nf.pack(fill=tk.X)
        _lbl(nf,"Client NDA status:",fnt=FONT_B).pack(side=tk.LEFT,padx=(0,10))
        tk.Checkbutton(nf,text="NDA Signed",variable=self.nda_signed,font=FONT_N,
                       bg=BH_WHITE,selectcolor=BH_WHITE,activebackground=BH_WHITE,
                       command=self._update_nda).pack(side=tk.LEFT,padx=6)
        self._nda_lbl=_lbl(nf,"✗  Not Signed",fnt=FONT_B,fg=BH_RED)
        self._nda_lbl.pack(side=tk.LEFT,padx=10)
        self._nda_note=_lbl(f,"",fnt=FONT_SM,fg=BH_GRAY)
        self._nda_note.pack(anchor=tk.W,pady=(8,0))

    def _kv(self, g, r, c, lbl, var, w=28, cs=1):
        _lbl(g,f"{lbl}:",fnt=FONT_B).grid(row=r,column=c,sticky=tk.W,padx=(0,6),pady=5)
        e=_entry(g,var,w=w)
        e.grid(row=r,column=c+1,columnspan=cs,sticky=tk.W,padx=(0,20),pady=5)
        self._w_by_var[str(var)]=e

    def _contact(self, p, col, title, nv, pv, ev, role=None):
        b=tk.Frame(p,bg=BH_WHITE); b.grid(row=0,column=col,sticky=tk.NW,padx=4)
        _lbl(b,title,fnt=FONT_S,fg=BH_GREEN).grid(row=0,column=0,columnspan=2,sticky=tk.W,pady=(0,8))
        for i,(l,v) in enumerate([("Name",nv),("Phone",pv),("Email",ev)],1):
            _lbl(b,f"{l}:",fnt=FONT_B).grid(row=i,column=0,sticky=tk.W,pady=4,padx=(0,8))
            if l == "Name" and role:
                cb = ttk.Combobox(b, textvariable=v, values=[], width=25,
                                  font=FONT_N, style="BH.TCombobox")
                cb.grid(row=i,column=1,sticky=tk.W,pady=4)
                cb.bind("<<ComboboxSelected>>", lambda e, r=role: self._contact_pick(r))
                cb.bind("<FocusOut>", lambda e, r=role: self._contact_pick(r, typed=True))
                cb.bind("<Return>", lambda e, r=role: self._contact_pick(r, typed=True))
                self._contact_boxes[role] = cb
                self._w_by_var[str(v)] = cb
            else:
                _entry(b,v,w=27).grid(row=i,column=1,sticky=tk.W,pady=4)

    # ── contact list ──────────────────────────────────────────────────────────
    ROLE_VARS = {"sales": ("sales_name", "sales_phone", "sales_email", "sales_title"),
                 "adv":   ("adv_name", "adv_phone", "adv_email", "adv_title"),
                 "ae":    ("ae_name", "ae_phone", "ae_email", "ae_title")}

    def _contacts_load(self, quiet=False, path=None):
        path = path or find_contacts_file(_tool_dir(), self.contacts_file.get())
        self.contacts = {}
        err = ""
        if path:
            try:
                self.contacts = read_contacts(path)
            except Exception as e:
                err = str(e)
        self.contacts_file.set(path or "")
        for role, cb in getattr(self, "_contact_boxes", {}).items():
            cb.config(values=[c["name"] for c in self.contacts.get(role, [])])
        self._contacts_paint(err)
        if not quiet and path:
            n = sum(len(v) for v in self.contacts.values())
            self._status.set(f"Contact list read — {n} people.")

    def _contacts_paint(self, err=""):
        if not hasattr(self, "_contacts_lbl"):
            return
        for w in self._contacts_btns.winfo_children():
            w.destroy()
        path = self.contacts_file.get()
        if path and not err:
            n = {r: len(v) for r, v in self.contacts.items()}
            self._contacts_lbl.config(
                text=f"Contact list:  {os.path.basename(path)}      "
                     f"{n.get('sales', 0)} sales managers  ·  "
                     f"{n.get('adv', 0)} sales advisors  ·  "
                     f"{n.get('ae', 0)} application engineers", fg=BH_GRAY)
            _ghost_btn(self._contacts_btns, "Open",
                       lambda: self._open_path(path)).pack(side=tk.RIGHT)
            _ghost_btn(self._contacts_btns, "Reload",
                       lambda: self._contacts_load()).pack(side=tk.RIGHT)
        else:
            self._contacts_lbl.config(
                text=(f"The contact list could not be read: {err}" if err else
                      "No contact list next to the tool. Names are typed in by hand."),
                fg=BH_WARN if err else BH_GRAY)
            _ghost_btn(self._contacts_btns, "Create the list",
                       self._contacts_create).pack(side=tk.RIGHT)
        _ghost_btn(self._contacts_btns, "Choose…",
                   self._contacts_choose, fg=BH_GRAY).pack(side=tk.RIGHT)

    def _contacts_choose(self):
        p = filedialog.askopenfilename(
            title="Contact list", initialdir=_tool_dir(),
            filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.contacts_file.set(p)
            self._contacts_load(path=p)
            self._save_settings()

    def _contacts_create(self):
        p = os.path.join(_tool_dir(), "Contacts.xlsx")
        if not os.path.exists(p):
            try:
                write_contacts_template(p)
            except Exception as e:
                messagebox.showerror("Contact list", str(e)); return
        self.contacts_file.set(p)
        self._contacts_load(quiet=True)
        self._open_path(p)

    def _contact_pick(self, role, typed=False):
        """A name from the list fills phone, email and job title."""
        names = self.ROLE_VARS[role]
        name = getattr(self, names[0]).get().strip()
        rec = next((c for c in self.contacts.get(role, [])
                    if c["name"].strip().lower() == name.lower()), None)
        if not rec:
            if not typed:
                return
            getattr(self, names[3]).set("")
            return
        getattr(self, names[0]).set(rec["name"])
        if rec.get("phone") or not typed:
            getattr(self, names[1]).set(rec.get("phone", ""))
        if rec.get("email") or not typed:
            getattr(self, names[2]).set(rec.get("email", ""))
        getattr(self, names[3]).set(rec.get("title", ""))

    def _contacts_defaults(self):
        """Fill the empty contacts: the person used last time, the
        application engineer matching the Windows login, or the only person
        of the list."""
        last = getattr(self, "_last_contacts", {}) or {}
        tokens = login_tokens()
        for role, people in self.contacts.items():
            var = getattr(self, self.ROLE_VARS[role][0])
            if var.get().strip() or not people:
                continue
            pick = next((c for c in people if c["name"] == last.get(role)), None)
            if pick is None and role == "ae":
                scored = sorted(((contact_matches_login(c, tokens), i)
                                 for i, c in enumerate(people)), reverse=True)
                if scored and scored[0][0] > 0:
                    pick = people[scored[0][1]]
            if pick is None and len(people) == 1:
                pick = people[0]
            if pick is not None:
                var.set(pick["name"])
                self._contact_pick(role)

    # ── TAB 2: TECHNICAL ──────────────────────────────────────────────────────
    def _tab_technical(self):
        p = self.t_technical

        f = self._section(p, "Compressor Configuration")
        g = tk.Frame(f,bg=BH_WHITE); g.pack(fill=tk.X)
        _lbl(g,"No. of Impellers:",fnt=FONT_B).grid(row=0,column=0,sticky=tk.W,padx=(0,6),pady=6)
        e=_combo(g,self.num_impellers,[str(i) for i in range(1,10)],w=5)
        e.grid(row=0,column=1,sticky=tk.W,padx=(0,20),pady=6)
        self._w_by_var[str(self.num_impellers)]=e
        _lbl(g,"Casing Size:",fnt=FONT_B).grid(row=0,column=2,sticky=tk.W,padx=(0,6),pady=6)
        _combo(g,self.casing_size,CASINGS,w=7).grid(
            row=0,column=3,sticky=tk.W,padx=(0,20),pady=6)
        _lbl(g,"Phases:",fnt=FONT_B).grid(row=0,column=4,sticky=tk.W,padx=(0,6))
        pc=_combo(g,self.num_phases,["1","2"],w=5)
        pc.grid(row=0,column=5,sticky=tk.W,padx=(0,20))
        pc.bind("<<ComboboxSelected>>",self._on_phase)
        _lbl(g,"Configuration Code:",fnt=FONT_B).grid(row=1,column=0,columnspan=2,
                                                      sticky=tk.W,pady=10)
        ce = tk.Entry(g,textvariable=self.config_code,width=20,font=FONT_C,
                      bg=BH_SEC,fg=BH_GREEN,insertbackground=BH_GREEN,
                      relief=tk.FLAT,bd=0,highlightthickness=1,
                      highlightbackground=BH_BORDER,highlightcolor=BH_GREEN)
        ce.grid(row=1,column=2,columnspan=2,sticky=tk.W,padx=(0,20))
        self._code_entry = ce
        Tooltip(ce, "Built from impellers, casing and phases. Type over it to "
                    "force another name (e.g. XICL455); it is then kept as typed.")
        hint = tk.Frame(g, bg=BH_WHITE)
        hint.grid(row=1,column=4,columnspan=4,sticky=tk.W)
        self._code_hint = _lbl(hint,"",fnt=FONT_SM,fg=BH_MUTED)
        self._code_hint.pack(side=tk.LEFT)
        self._code_reset = _ghost_btn(hint, "Back to automatic",
                                      self._code_to_auto, fg=BH_GREEN)
        self.config_code.trace_add("write", self._code_typed)
        self._code_hint_paint()

        f = self._section(p, "Equipment")
        g = tk.Frame(f,bg=BH_WHITE); g.pack(fill=tk.X)
        for ci,(lbl,var,vals,idx) in enumerate([
            ("Motor Type",    self.motor_type,  ["MGV0","MGV1","MGV2","MGV3"],1),
            ("Configuration", self.back_to_back,["Back to Back","Simple"],   0),
            ("Rating (ANSI)", self.rating,       RATINGS,                    1),
        ]):
            _lbl(g,f"{lbl}:",fnt=FONT_B).grid(row=0,column=ci*2,sticky=tk.W,padx=(0,6),pady=6)
            cb=_combo(g,var,vals,w=14)
            cb.grid(row=0,column=ci*2+1,sticky=tk.W,padx=(0,20),pady=6); cb.current(idx)
        _lbl(g,"UCP / PLC Type:",fnt=FONT_B).grid(row=1,column=0,sticky=tk.W,padx=(0,6),pady=6)
        plc_cb=_combo(g,self.ucp_type,list(PLC_DOCS.keys()),w=30)
        plc_cb.grid(row=1,column=1,columnspan=3,sticky=tk.W,padx=(0,20),pady=6)
        plc_cb.bind("<<ComboboxSelected>>",self._on_plc)
        _lbl(g,"PLC Option:",fnt=FONT_B).grid(row=1,column=4,sticky=tk.W,padx=(0,6),pady=6)
        self._plc_opt=_combo(g,self.plc_option,
                             ["Base Case","Intrinsically Safe","Surge Arrestor"],w=20)
        self._plc_opt.grid(row=1,column=5,sticky=tk.W,padx=(0,20),pady=6)
        self._plc_opt.bind("<<ComboboxSelected>>",self._on_plc)
        self._plc_badge=tk.Frame(g,bg=BH_SEC,highlightbackground=BH_BORDER,highlightthickness=1)
        self._plc_badge.grid(row=2,column=0,columnspan=8,sticky=tk.EW,pady=8,ipady=6,ipadx=10)
        _lbl(self._plc_badge,"Associated PLC Documents:",fnt=FONT_B,
             bg=BH_SEC,fg=BH_GREEN).pack(anchor=tk.W,padx=8,pady=(4,2))
        self._plc_doc_lbl=_lbl(self._plc_badge,"",fnt=FONT_C,bg=BH_SEC)
        self._plc_doc_lbl.pack(anchor=tk.W,padx=8,pady=(0,4))
        tk.Checkbutton(g,text="Safety & Control PLC Included",variable=self.plc_safety,
                       font=FONT_N,bg=BH_WHITE,selectcolor=BH_WHITE,
                       activebackground=BH_WHITE).grid(row=3,column=0,columnspan=2,sticky=tk.W,pady=4)
        tk.Checkbutton(g,text="Thrust Bearing Included",variable=self.thrust_bearing,
                       font=FONT_N,bg=BH_WHITE,selectcolor=BH_WHITE,
                       activebackground=BH_WHITE).grid(row=3,column=2,columnspan=2,sticky=tk.W,pady=4)

        f = self._section(p, "Cooling Systems")
        g = tk.Frame(f,bg=BH_WHITE); g.pack(fill=tk.X)
        _lbl(g,"Cooling System Concept:",fnt=FONT_B,fg=BH_GREEN).grid(
            row=0,column=0,columnspan=2,sticky=tk.W,pady=4)
        e=_combo(g,self.cooling_system,COOLING_CONCEPTS,w=52)
        e.grid(row=0,column=2,columnspan=4,sticky=tk.W,pady=4)
        self._w_by_var[str(self.cooling_system)]=e
        _lbl(g,"VFD Cooling:",fnt=FONT_B).grid(row=1,column=0,sticky=tk.W,padx=(0,6),pady=6)
        _combo(g,self.vfd_cooling,["Cooler","Chiller"],w=12).grid(
            row=1,column=1,sticky=tk.W,padx=(0,20))
        _lbl(g,"Transformer Type:",fnt=FONT_B).grid(row=2,column=0,sticky=tk.W,padx=(0,6),pady=6)
        _combo(g,self.transfo_type,["Oil ONAN","Oil ONAF","Dry Type","Dry Zone"],w=16).grid(
            row=2,column=1,sticky=tk.W,padx=(0,20))
        _lbl(g,"Required Power (kW):",fnt=FONT_B).grid(row=2,column=2,sticky=tk.W,padx=(20,6))
        _entry(g,self.power,w=14).grid(row=2,column=3,sticky=tk.W,padx=(0,20))
        # grid-compatible separator

        f = self._section(p, "Documents for this configuration",
                          "picked in the document library, updated as you type")
        g = tk.Frame(f,bg=BH_WHITE); g.pack(fill=tk.X)
        g.columnconfigure(1, weight=1)
        self._tech_prev = {}
        for r,(lid,label) in enumerate((("cooling","Cooling concept"),
                                        ("pid","P&ID"),("gad","GAD"),
                                        ("plc","PLC documents"))):
            _lbl(g,label,fnt=FONT_B).grid(row=r,column=0,sticky=tk.W,padx=(0,18),pady=4)
            fl=_lbl(g,"—",fnt=FONT_N,fg=BH_INK,anchor=tk.W,justify=tk.LEFT,
                    wraplength=760)
            fl.grid(row=r,column=1,sticky=tk.W,pady=4)
            ch=tk.Frame(g,bg=BH_WHITE); ch.grid(row=r,column=2,sticky=tk.E,padx=(12,0))
            self._tech_prev[lid]=(fl,ch)

        f = self._section(p, "Detailed Scope")
        g = tk.Frame(f,bg=BH_WHITE); g.pack(fill=tk.X)
        tk.Checkbutton(g,text="Antisurge Valve",variable=self.antisurge_valve,
                       font=FONT_N,bg=BH_WHITE,selectcolor=BH_WHITE,
                       activebackground=BH_WHITE).grid(row=0,column=0,sticky=tk.W,pady=4)
        _lbl(g,"Filter Type:",fnt=FONT_B).grid(row=0,column=1,sticky=tk.W,padx=(20,6))
        _combo(g,self.filter_type,["Simplex","Duplex"],w=12).grid(
            row=0,column=2,sticky=tk.W,padx=(0,20))
        _lbl(g,"Extra Scope:",fnt=FONT_B,fg=BH_GREEN).grid(
            row=1,column=0,columnspan=3,sticky=tk.W,pady=(10,4))
        for ci,(txt,var) in enumerate([("E-House",self.ehouse),("MCC",self.mcc),
                                        ("Compressor Enclosure",self.comp_enclosure)]):
            tk.Checkbutton(g,text=txt,variable=var,font=FONT_N,bg=BH_WHITE,
                           selectcolor=BH_WHITE,activebackground=BH_WHITE).grid(
                row=2,column=ci,sticky=tk.W,padx=12)

        f = self._section(p, "General Clarifications & Notes")
        tf=tk.Frame(f,bg=BH_WHITE); tf.pack(fill=tk.X)
        self._clarif=tk.Text(tf,height=5,width=120,font=FONT_N,wrap=tk.WORD,
                             relief=tk.FLAT,bd=0,highlightthickness=1,
                             highlightbackground=BH_BORDER,highlightcolor=BH_GREEN)
        self._clarif.pack(side=tk.LEFT,fill=tk.X,expand=True)
        sb=ttk.Scrollbar(tf,command=self._clarif.yview)
        sb.pack(side=tk.RIGHT,fill=tk.Y)
        self._clarif.config(yscrollcommand=sb.set)

    # ── TAB 3: DOCUMENTS ──────────────────────────────────────────────────────
    QUALITY = {"exact":   ("Exact", BH_OK),
               "chosen":  ("Chosen", BH_OK),
               "project": ("Project folder", BH_OK),
               "closest": ("Closest", BH_WARN),
               "tool":    ("Built by the tool", BH_TOOL),
               "missing": ("Not found", BH_RED)}

    def _tab_docs(self):
        p = self.t_docs
        f = self._section(p, "Documents to provide",
                          "ticked documents are placed in the proposal")
        top = tk.Frame(f, bg=BH_WHITE); top.pack(fill=tk.X, pady=(0, 10))
        self._lib_lbl = _lbl(top, "", fnt=FONT_SM, fg=BH_GRAY)
        self._lib_lbl.pack(side=tk.LEFT)
        _btn(top, "Defaults", self._lib_defaults, bg=BH_GRAY).pack(side=tk.RIGHT)
        _btn(top, "Open the library", lambda: self._open_path(self.db_root.get()),
             bg=BH_GRAY).pack(side=tk.RIGHT, padx=8)
        Tooltip(top.winfo_children()[-2],
                "Back to the documents ticked by default for this offer category")

        hdr = tk.Frame(f, bg=BH_GREEN); hdr.pack(fill=tk.X)
        tk.Frame(hdr, bg=BH_GREEN, width=32, height=1).pack(side=tk.LEFT)
        tk.Label(hdr, text="Document", font=FONT_B, bg=BH_GREEN, fg=BH_WHITE,
                 width=30, anchor=tk.W, pady=5).pack(side=tk.LEFT)
        tk.Label(hdr, text="Fit", font=FONT_B, bg=BH_GREEN, fg=BH_WHITE,
                 width=16).pack(side=tk.RIGHT, padx=(0, 70))
        tk.Label(hdr, text="File used", font=FONT_B, bg=BH_GREEN, fg=BH_WHITE,
                 anchor=tk.W).pack(side=tk.LEFT)

        body = tk.Frame(f, bg=BH_WHITE); body.pack(fill=tk.X)
        self._lib_rows = {}
        for i, doc in enumerate(DB_LIBRARY):
            bg = BH_WHITE if i % 2 == 0 else BH_LIGHT
            row = tk.Frame(body, bg=bg); row.pack(fill=tk.X)
            tk.Frame(row, bg=BH_BORDER, height=1).pack(fill=tk.X)
            inner = tk.Frame(row, bg=bg, pady=4); inner.pack(fill=tk.X)
            cb = tk.Checkbutton(inner, variable=self.lib_vars[doc["id"]], bg=bg,
                                selectcolor=BH_WHITE, activebackground=bg, bd=0,
                                highlightthickness=0)
            cb.pack(side=tk.LEFT, padx=(6, 6))
            names = tk.Frame(inner, bg=bg, width=230)
            names.pack(side=tk.LEFT);
            nm = _lbl(names, doc["name"], fnt=FONT_B, bg=bg, width=30, anchor=tk.W)
            nm.pack(anchor=tk.W)
            _lbl(names, doc["title"], fnt=FONT_SM, fg=BH_MUTED, bg=bg, width=34,
                 anchor=tk.W).pack(anchor=tk.W)
            act = tk.Frame(inner, bg=bg, width=70); act.pack(side=tk.RIGHT, padx=(0, 6))
            chip = tk.Frame(inner, bg=bg, width=130); chip.pack(side=tk.RIGHT, padx=8)
            fl = _lbl(inner, "", fnt=FONT_N, bg=bg, anchor=tk.W, justify=tk.LEFT,
                      wraplength=560)
            fl.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self._lib_rows[doc["id"]] = dict(file=fl, chip=chip, act=act, bg=bg,
                                             name=nm)
        _lbl(f, "Unticked documents stay out of the proposal and out of its "
                "numbering. A file picked with “Change” is kept until "
                "“Automatic”.", fnt=FONT_SM, fg=BH_MUTED).pack(anchor=tk.W,
                                                                 pady=(10, 0))

    # ── library ──────────────────────────────────────────────────────────────
    def _lib_on(self, lid):
        v = self.lib_vars.get(lid)
        return bool(v.get()) if v is not None else False

    def _lib_toggled(self, lid):
        if self._lib_loading:
            return
        self.lib_state.setdefault(self.offer_category.get(), {})[lid] = \
            self._lib_on(lid)
        self._struct_cache = (None, None)
        self._struct_cache2 = {}
        self._lib_paint()
        self._refresh_docs()

    def _lib_load_category(self, cat):
        """The ticks follow the offer category: the defaults of the category,
        changed by what was ticked by hand for it."""
        state = self.lib_state.get(cat, {})
        self._lib_loading = True
        try:
            for doc in DB_LIBRARY:
                self.lib_vars[doc["id"]].set(state.get(doc["id"], doc["on"][cat]))
        finally:
            self._lib_loading = False
        self._lib_cat = cat
        self._struct_cache = (None, None)
        self._struct_cache2 = {}

    def _lib_defaults(self):
        cat = self.offer_category.get()
        self.lib_state.pop(cat, None)
        self._lib_load_category(cat)
        self._lib_paint()
        self._refresh_docs()

    def _lib_ctx(self, project_files=()):
        return dict(concept=self.cooling_system.get(),
                    phases=self.num_phases.get().strip(),
                    casing=_casing_digits(self.casing_size.get()),
                    impellers=self.num_impellers.get().strip(),
                    motor=self.motor_type.get().strip(),
                    ucp=self.ucp_type.get(), plc_option=self.plc_option.get(),
                    project_files=list(project_files))

    def _lib_update(self, project_files=()):
        ctx = self._lib_ctx(project_files)
        self._lib_res = {}
        for doc in DB_LIBRARY:
            try:
                self._lib_res[doc["id"]] = lib_resolve(
                    doc, self._lib_folders.get(doc["id"]), ctx,
                    self.lib_choice.get(doc["id"]))
            except Exception as e:
                self._lib_res[doc["id"]] = dict(files=[], quality="missing",
                                                note=f"error: {e}", candidates=[])
        self._lib_paint()

    def _lib_chip(self, parent, res, bg):
        for w in parent.winfo_children():
            w.destroy()
        txt, col = self.QUALITY.get(res.get("quality"), ("—", BH_GRAY))
        _pill(parent, txt, col, host_bg=bg).pack(anchor=tk.E)

    def _lib_paint(self):
        if not hasattr(self, "_lib_rows"):
            return
        root = self.db_root.get().strip()
        if root and os.path.isdir(root):
            self._lib_lbl.config(
                text=f"Library:  {root}      {len(self._lib_folders)} of "
                     f"{len(DB_LIBRARY)} document folders found", fg=BH_GRAY)
        else:
            self._lib_lbl.config(text="Library folder not found: put DB-2026 next "
                                      "to the tool, or pick it in the document "
                                      "manager.", fg=BH_WARN)
        for lid, w in self._lib_rows.items():
            res = self._lib_res.get(lid, {})
            on = self._lib_on(lid)
            files = res.get("files") or []
            sig = (on, tuple(files), res.get("quality"), res.get("note"),
                   len(res.get("candidates") or []))
            if w.get("_sig") == sig:
                continue
            w["_sig"] = sig
            text = ",  ".join(os.path.basename(x) for x in files[:3])
            if len(files) > 3:
                text += f"   +{len(files) - 3}"
            if res.get("note") and res.get("quality") not in ("exact", "tool"):
                text = (text + "\n" if text else "") + res["note"]
            elif res.get("quality") == "tool":
                text = "generated from the scope of supply page"
            w["file"].config(text=text or "—", fg=BH_INK if on else BH_MUTED)
            w["name"].config(fg=BH_INK if on else BH_MUTED)
            if on:
                self._lib_chip(w["chip"], res, w["bg"])
            else:
                for c in w["chip"].winfo_children():
                    c.destroy()
                _lbl(w["chip"], "not in this offer", fnt=FONT_SM, fg=BH_MUTED,
                     bg=w["bg"]).pack(anchor=tk.E)
            for c in w["act"].winfo_children():
                c.destroy()
            if len(res.get("candidates") or []) > 1 and \
                    DB_BY_ID[lid]["rule"] not in ("all",):
                _ghost_btn(w["act"], "Change",
                           lambda e=None, i=lid, a=w["act"]: self._lib_menu(i, a),
                           bg=w["bg"]).pack(side=tk.RIGHT)
        for lid, (fl, ch) in getattr(self, "_tech_prev", {}).items():
            res = self._lib_res.get(lid, {})
            files = res.get("files") or []
            text = ",  ".join(os.path.basename(x) for x in files[:4]) or "—"
            if res.get("note") and res.get("quality") not in ("exact", "tool"):
                text += f"\n{res['note']}"
            fl.config(text=text)
            if getattr(ch, "_q", None) != res.get("quality"):
                ch._q = res.get("quality")
                self._lib_chip(ch, res, BH_WHITE)

    def _lib_menu(self, lid, anchor=None, event=None):
        res = self._lib_res.get(lid, {})
        m = tk.Menu(self.root, tearoff=0, bg=BH_WHITE, fg=BH_INK,
                    activebackground=BH_GREEN, activeforeground=BH_WHITE,
                    font=FONT_N, bd=0)
        cur = (res.get("files") or [None])[0]
        for path, why in (res.get("candidates") or [])[:25]:
            label = os.path.basename(path) + (f"     {why}" if why else "")
            m.add_command(label=("●  " if path == cur else "     ") + label,
                          command=lambda p=path: self._lib_pick(lid, p))
        m.add_separator()
        m.add_command(label="Automatic", command=lambda: self._lib_pick(lid, None))
        if res.get("folder"):
            m.add_command(label="Open the folder",
                          command=lambda: self._open_path(res["folder"]))
        try:
            if event is not None:
                m.tk_popup(event.x_root, event.y_root)
            else:
                m.tk_popup(anchor.winfo_rootx(),
                           anchor.winfo_rooty() + anchor.winfo_height())
        finally:
            m.grab_release()

    def _lib_pick(self, lid, path):
        if path:
            self.lib_choice[lid] = path
        else:
            self.lib_choice.pop(lid, None)
        self._scan_docs(silent=True, rewalk=False)

    def _add_doc(self):
        rf=tk.Frame(self._extra_doc_frame,bg=BH_WHITE); rf.pack(fill=tk.X,pady=2)
        chk=tk.BooleanVar(value=True); nv=tk.StringVar()
        tk.Checkbutton(rf,variable=chk,bg=BH_WHITE,selectcolor=BH_WHITE,
                       activebackground=BH_WHITE).pack(side=tk.LEFT)
        _entry(rf,nv,w=42).pack(side=tk.LEFT,padx=6)
        self.extra_docs.append((nv,chk))
        def _rm(f=rf,n=nv):
            self.extra_docs=[(a,b) for a,b in self.extra_docs if a is not n]; f.destroy()
        _btn(rf,"✕",_rm,bg=BH_RED,fnt=FONT_SM).pack(side=tk.LEFT,padx=4)

    def _refresh_plc_container(self):
        for w in self._plc_doc_container.winfo_children(): w.destroy()
        plc=self.ucp_type.get(); opt=self.plc_option.get()
        docs=PLC_DOCS.get(plc,{}).get(opt,[])
        if docs:
            _lbl(self._plc_doc_container,f"  {plc}  |  Option: {opt}",
                 fnt=FONT_B,fg=BH_GREEN).pack(anchor=tk.W,pady=(0,4))
            rw=tk.Frame(self._plc_doc_container,bg=BH_SEC,padx=10,pady=6)
            rw.pack(fill=tk.X)
            for d in docs:
                tk.Label(rw,text=d,font=FONT_C,bg=BH_SEC,fg=BH_DARK,
                         padx=8,pady=2).pack(side=tk.LEFT,padx=4)
        else:
            _lbl(self._plc_doc_container,"No PLC documents for this selection.",
                 fnt=FONT_N,fg=BH_GRAY).pack(anchor=tk.W)

    # ── PAGE 4: DOCUMENT MANAGER ──────────────────────────────────────────────
    def _tab_docmgr(self):
        p = self.t_docmgr
        self._pdf_hero(p)

        # ---- sources -----------------------------------------------------
        f = self._section(p, "Document Sources",
                          "project folder  ·  local standard database")
        g = tk.Frame(f, bg=BH_WHITE); g.pack(fill=tk.X)
        g.columnconfigure(1, weight=1)

        _lbl(g, "Project folder", fnt=FONT_B).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 14), pady=6)
        pe = _entry(g, self.user_root, w=60)
        self._w_by_var[str(self.user_root)] = pe
        pe.grid(row=0, column=1, sticky=tk.EW, padx=(0, 12), pady=6, ipady=4)
        Tooltip(pe, "Folder holding the customer / project documents.\n"
                    "Layout: Section 1_Executive Summary, Section 2_Definition of supply, …")
        b1 = tk.Frame(g, bg=BH_WHITE); b1.grid(row=0, column=2, sticky=tk.W)
        _btn(b1, "Browse", self._pick_user_root, bg=BH_GRAY).pack(side=tk.LEFT, padx=2)
        _btn(b1, "Create structure", self._create_structure).pack(side=tk.LEFT, padx=2)
        _btn(b1, "Open", lambda: self._open_path(self.user_root.get()),
             bg=BH_GRAY).pack(side=tk.LEFT, padx=2)

        _lbl(g, "Document library", fnt=FONT_B).grid(
            row=1, column=0, sticky=tk.W, padx=(0, 14), pady=6)
        de = _entry(g, self.db_root, w=60)
        de.grid(row=1, column=1, sticky=tk.EW, padx=(0, 12), pady=6, ipady=4)
        Tooltip(de, "DB-2026: one folder per document (GAD, P&ID, Cooling_Concept, "
                    "QCP…). Found on its own next to the tool.")
        b2 = tk.Frame(g, bg=BH_WHITE); b2.grid(row=1, column=2, sticky=tk.W)
        _btn(b2, "Browse", self._pick_db_root, bg=BH_GRAY).pack(side=tk.LEFT, padx=2)
        _btn(b2, "Initialise", self._init_db).pack(side=tk.LEFT, padx=2)
        _btn(b2, "Open", lambda: self._open_path(self.db_root.get()),
             bg=BH_GRAY).pack(side=tk.LEFT, padx=2)

        opt = tk.Frame(f, bg=BH_SEC, padx=14, pady=10,
                       highlightbackground=BH_BORDER, highlightthickness=1)
        opt.pack(fill=tk.X, pady=(14, 0))
        tk.Checkbutton(opt, text="  Use standard documents from the local database",
                       variable=self.use_db, font=FONT_B, bg=BH_SEC, fg=BH_INK,
                       selectcolor=BH_WHITE, activebackground=BH_SEC, bd=0,
                       highlightthickness=0,
                       command=self._refresh_docs).pack(side=tk.LEFT)
        _lbl(opt, "recommended for DTI and Budget offers", fnt=FONT_SM,
             fg=BH_GRAY, bg=BH_SEC).pack(side=tk.LEFT, padx=(8, 28))
        tk.Checkbutton(opt, text="  Copy them into the project folder",
                       variable=self.auto_place, font=FONT_N, bg=BH_SEC,
                       fg=BH_INK, selectcolor=BH_WHITE, bd=0, highlightthickness=0,
                       activebackground=BH_SEC).pack(side=tk.LEFT)

        # ---- structure ---------------------------------------------------
        f2 = self._section(p, "Technical Proposal Structure",
                           "every section and sub-section of the generated offer")

        bar = tk.Frame(f2, bg=BH_WHITE); bar.pack(fill=tk.X, pady=(0, 4))
        _btn(bar, "Scan documents", self._scan_docs).pack(side=tk.LEFT)
        _btn(bar, "Fill from database", self._fill_from_db,
             bg=BH_TOOL).pack(side=tk.LEFT, padx=8)
        _btn(bar, "Expand all", lambda: self._collapse_all(False),
             bg=BH_GRAY).pack(side=tk.LEFT)
        _btn(bar, "Collapse all", lambda: self._collapse_all(True),
             bg=BH_GRAY).pack(side=tk.LEFT, padx=8)
        _btn(bar, "Add section", self._add_section, bg=BH_GRAY).pack(side=tk.LEFT)
        rb = _btn(bar, "Restore default", self._reset_order, bg=BH_GRAY)
        rb.pack(side=tk.LEFT, padx=8)
        self._off_toggle = tk.Label(bar, text="", font=FONT_SM, bg=BH_WHITE,
                                    fg=BH_GREEN, cursor="hand2")
        self._off_toggle.pack(side=tk.LEFT, padx=10)
        self._off_toggle.bind("<Button-1>", lambda e: self._toggle_off())
        Tooltip(rb, "Back to the standard sections, lines and order")

        sb = tk.Frame(bar, bg=BH_LIGHT, highlightbackground=BH_BORDER,
                      highlightthickness=1)
        sb.pack(side=tk.RIGHT)
        _lbl(sb, "Filter", fnt=FONT_BD, bg=BH_LIGHT, fg=BH_GRAY).pack(side=tk.LEFT, padx=(10, 6))
        se = self._doc_search_entry = tk.Entry(sb, textvariable=self.doc_search, width=24, font=FONT_N,
                      relief=tk.FLAT, bd=0, bg=BH_WHITE, fg=BH_INK,
                      insertbackground=BH_GREEN)
        se.pack(side=tk.LEFT, padx=(0, 2), pady=5, ipady=2)
        _ghost_btn(sb, "✕", lambda: (self.doc_search.set(""), self._apply_visibility()),
                   fg=BH_GRAY, bg=BH_LIGHT).pack(side=tk.LEFT, padx=(0, 6))
        se.bind("<KeyRelease>", lambda e: self._apply_visibility())
        Tooltip(se, "Filter by document name or number")

        # filter chips
        fl = tk.Frame(f2, bg=BH_WHITE); fl.pack(fill=tk.X, pady=(0, 10))
        self._filter_btns = {}
        self._filter_txt = {"all": "All", "todo": "To provide",
                            "db": "From database", "ok": "Available"}
        for key, txt in self._filter_txt.items():
            b = tk.Label(fl, text=f"  {txt}  ", font=FONT_BD, bg=BH_LIGHT,
                         fg=BH_GRAY, padx=6, pady=4, cursor="hand2")
            b.pack(side=tk.LEFT, padx=(0, 6))
            b.bind("<Button-1>", lambda e, k=key: self._set_filter(k))
            self._filter_btns[key] = b

        nav = tk.Frame(f2, bg=BH_WHITE)
        nav.pack(fill=tk.X, pady=(0, 8))
        _lbl(nav, "Jump to", fnt=FONT_XS, fg=BH_MUTED).pack(side=tk.LEFT, padx=(0, 8))
        self._nav_holder = tk.Frame(nav, bg=BH_WHITE)
        self._nav_holder.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._nav_chips = {}

        self._doc_container = tk.Frame(f2, bg=BH_WHITE)
        self._doc_container.pack(fill=tk.X)
        self._doc_empty = tk.Frame(f2, bg=BH_LIGHT, pady=26)
        _lbl(self._doc_empty, "Nothing matches this filter.", fnt=FONT_S,
             fg=BH_GRAY, bg=BH_LIGHT).pack()
        _lbl(self._doc_empty, "Clear the search box or switch back to “All”.",
             fnt=FONT_SM, fg=BH_MUTED, bg=BH_LIGHT).pack(pady=(4, 0))

        se.bind("<Escape>", lambda e: (self.doc_search.set(""),
                                       self._apply_visibility()))
        self.root.bind("<Control-f>", lambda e: (self._show_page(5), se.focus_set()))

        sm = tk.Frame(f2, bg=BH_WHITE); sm.pack(fill=tk.X, pady=(14, 0))
        self._doc_progress = ttk.Progressbar(sm, style="BH.Horizontal.TProgressbar",
                                             mode="determinate", maximum=100)
        self._doc_progress.pack(fill=tk.X, pady=(0, 7))
        self._doc_summary = tk.StringVar(value="No scan yet.")
        tk.Label(sm, textvariable=self._doc_summary, font=FONT_B,
                 bg=BH_WHITE, fg=BH_GRAY, anchor=tk.W,
                 justify=tk.LEFT, wraplength=1200).pack(anchor=tk.W)

        self._set_filter("all", render=False)

    # ── the big generate button ───────────────────────────────────────────────
    def _pdf_hero(self, parent):
        """The one action the whole tool exists for, given its weight without
        decoration: dark panel, single brand rule, one solid button."""
        shell = tk.Frame(parent, bg=BH_BG)
        shell.pack(fill=tk.X, padx=SP4, pady=(0, SP3))
        frame = tk.Frame(shell, bg=BH_NAVY)
        frame.pack(fill=tk.X)
        tk.Frame(frame, bg=BH_GREEN, width=3).pack(side=tk.LEFT, fill=tk.Y)
        card = tk.Frame(frame, bg=BH_INK, padx=SP5, pady=SP4)
        card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Frame(shell, bg=BH_SHADOW_2, height=1).pack(fill=tk.X)

        right = tk.Frame(card, bg=BH_INK)
        right.pack(side=tk.RIGHT, padx=(SP5, 0))
        b = _SoftButton(right, "Generate", self._generate_pdf,
                        fill=BH_GREEN, host_bg=BH_INK, fg=BH_WHITE,
                        fnt=_fnt(11, True), pad=(24, 12), radius=R_SM)
        b.pack()
        self._gen_buttons = getattr(self, "_gen_buttons", []) + [b]
        Tooltip(b, "Build the full proposal: cover, contents, dividers, "
                   "every document in order   (Ctrl+G)")
        self._pdf_hint = tk.StringVar(value="")
        tk.Label(right, textvariable=self._pdf_hint, font=FONT_SM, bg=BH_INK,
                 fg=BH_ON_DARK).pack(pady=(9, 0))

        left = tk.Frame(card, bg=BH_INK)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        _lbl(left, "Generate the technical proposal", fnt=_fnt(14, True),
             fg=BH_WHITE, bg=BH_INK).pack(anchor=tk.W)
        _lbl(left, "Sections in the order set below, every document converted to "
                   "PDF and placed under its heading. With Technical_Proposal.docx "
                   "next to the tool, that Word file is the layout.",
             fnt=FONT_N, fg=BH_ON_DARK, bg=BH_INK, justify=tk.LEFT,
             wraplength=640).pack(anchor=tk.W, pady=(5, 10))

        self._ready_box = tk.Frame(left, bg=BH_INK)
        self._ready_box.pack(anchor=tk.W, fill=tk.X, pady=(0, 12))

        out = tk.Frame(left, bg=BH_INK)
        out.pack(anchor=tk.W, pady=(0, 10))
        _lbl(out, "Output", fnt=FONT_BD, fg=BH_ON_DARK, bg=BH_INK).pack(
            side=tk.LEFT, padx=(0, 10))
        self._out_btns = {}
        for key, txt, tip in [
                ("single", "One PDF", "The whole proposal in a single file."),
                ("folder", "One PDF per section",
                 "A folder holding one PDF for the front matter and one for "
                 "each section, numbered in order.")]:
            ob = tk.Label(out, text=f"  {txt}  ", font=FONT_BD, bg=BH_LINE_D,
                          fg=BH_ON_DARK2, padx=6, pady=4, cursor="hand2")
            ob.pack(side=tk.LEFT, padx=(0, 6))
            ob.bind("<Button-1>", lambda e, k=key: (self.pdf_output.set(k),
                                                    self._paint_output()))
            Tooltip(ob, tip)
            self._out_btns[key] = ob
        wc = tk.Checkbutton(out, text="  Word skeleton", variable=self.pdf_use_word,
                            font=FONT_SM, bg=BH_INK, fg=BH_ON_DARK2,
                            selectcolor=BH_LINE_D, activebackground=BH_INK,
                            activeforeground=BH_WHITE, bd=0, highlightthickness=0,
                            cursor="hand2")
        wc.pack(side=tk.LEFT, padx=(14, 0))
        Tooltip(wc, "Build from Technical_Proposal.docx kept in the same folder "
                    "as the tool. Without it, or when Word is not installed, "
                    "the built-in layout is used.")
        self._paint_output()

        sk = tk.Frame(left, bg=BH_INK)
        sk.pack(anchor=tk.W, pady=(0, 10))
        self._skel_lbl = _lbl(sk, "", fnt=FONT_SM, fg=BH_ON_DARK2, bg=BH_INK)
        self._skel_lbl.pack(side=tk.LEFT)
        ch = _lbl(sk, "   change", fnt=FONT_SM, fg=BH_ACCENT, bg=BH_INK,
                  cursor="hand2")
        ch.pack(side=tk.LEFT)
        ch.bind("<Button-1>", lambda e: self._pick_skeleton())
        self._paint_skeleton()

        opts = tk.Frame(left, bg=BH_INK)
        opts.pack(anchor=tk.W)
        for var, txt, tip in [
                (self.pdf_placeholders, "Placeholder for missing documents",
                 "Insert a “document to be provided” page so the proposal keeps "
                 "its full structure."),
                (self.pdf_stamp, "Headers & page numbers",
                 "Stamp customer name, Baker Hughes reference and “Page x of y” "
                 "on every page."),
                (self.pdf_scope, "Build the scope of supply",
                 "Print the BH / customer split from the Scope of supply page."),
                (self.pdf_commercial, "Build the commercial section",
                 "Price, INCOTERM, lead time and clarifications, for DTI and "
                 "Budget offers."),
                (self.pdf_skip_empty, "Drop empty sections",
                 "Leave out a section when none of its documents is available."),
                (self.pdf_frame, "Proposal header on inserted pages",
                 "Word skeleton only: A4 portrait documents are placed inside "
                 "the proposal page, under its header and above its footer. "
                 "Drawings and landscape sheets keep their own page.")]:
            cb = tk.Checkbutton(opts, text=f"  {txt}", variable=var, font=FONT_SM,
                                bg=BH_INK, fg=BH_ON_DARK2, selectcolor=BH_LINE_D,
                                activebackground=BH_INK, activeforeground=BH_WHITE,
                                bd=0, highlightthickness=0, cursor="hand2")
            cb.pack(side=tk.LEFT, padx=(0, 18))
            Tooltip(cb, tip)

    def _paint_output(self):
        for key, btn in getattr(self, "_out_btns", {}).items():
            on = key == self.pdf_output.get()
            btn.config(bg=BH_GREEN if on else BH_LINE_D,
                       fg=BH_WHITE if on else BH_ON_DARK2)

    def _skeleton_path(self):
        """The Word master kept next to the tool. Its exact name does not
        matter: a file picked by hand wins, then Technical_Proposal.docx,
        then any Word file whose name mentions the proposal, then the only
        Word file of the folder."""
        picked = self.skeleton_file.get().strip()
        if picked and os.path.exists(picked):
            return picked
        folder = _tool_dir()
        try:
            words = sorted(f for f in os.listdir(folder)
                           if f.lower().endswith((".docx", ".docm"))
                           and not f.startswith("~$"))
        except OSError:
            words = []
        low = {f.lower(): f for f in words}
        hit = low.get(SKELETON_NAME.lower())
        if not hit:
            named = [f for f in words if "proposal" in f.lower()]
            hit = named[0] if named else (words[0] if len(words) == 1 else None)
        return os.path.join(folder, hit) if hit else ""

    def _pick_skeleton(self):
        path = filedialog.askopenfilename(
            title="Word file used as the proposal master",
            initialdir=_tool_dir(),
            filetypes=[("Word document", "*.docx *.docm")])
        if path:
            self.skeleton_file.set(path)
            self._paint_skeleton()

    def _paint_skeleton(self):
        if not hasattr(self, "_skel_lbl"):
            return
        path = self._skeleton_path()
        if path:
            self._skel_lbl.config(text=f"Word master:  {os.path.basename(path)}",
                                  fg=BH_ON_DARK2)
        else:
            self._skel_lbl.config(
                text=f"No Word master found in {_tool_dir()}", fg=BH_WARN)

    def _prepare_docs(self, tmp, doc_status=None, nda=None, comm=None, say=None):
        """Document status with every file turned into PDF pages: Excel, Word
        and images are converted, several files under one heading are merged
        in the order they were picked. The commercial section of the project
        is filled with the offer values first; without a signed NDA the
        datasheet stops after the gas composition."""
        cache, out = {}, {}
        self._prep_report, self._commercial_copy = [], ""
        doc_status = self.doc_status if doc_status is None else doc_status
        nda = bool(self.nda_signed.get()) if nda is None else nda
        comm = self._commercial_values() if comm is None else comm
        todo = [c for c, st in doc_status.items() if st.get("paths") or st.get("path")]
        done = 0
        for code, st in doc_status.items():
            st2 = dict(st)
            paths = st.get("paths") or ([st["path"]] if st.get("path") else [])
            if paths and say:
                done += 1
                say(f"Converting {done} of {len(todo)} — "
                    f"{os.path.basename(paths[0])}")
            pdfs, notes = [], []
            for p in paths:
                src = p
                if st.get("commercial") and p.lower().endswith((".docx", ".docm")):
                    src = os.path.join(tmp, "Commercial_filled.docx")
                    try:
                        for field, what in fill_commercial_docx(
                                p, src, comm):
                            self._prep_report.append((st.get("ref", "11"),
                                                      "Commercial section", what))
                        self._commercial_copy = src
                    except Exception as e:
                        src = p
                        self._prep_report.append((st.get("ref", "11"),
                                                  "Commercial section",
                                                  f"NOT FILLED — {e}"))
                pdf, how = to_pdf(src, tmp, cache)
                if pdf:
                    pdfs.append(pdf)
                else:
                    notes.append(f"{os.path.basename(p)}: {how}")
            if st.get("nda") and pdfs and not nda:
                cut_pdfs = []
                for k, pdf in enumerate(pdfs):
                    res = cut_after_gas_composition(
                        pdf, os.path.join(tmp, f"nda_{k}_" + os.path.basename(pdf)))
                    if res is None:
                        st2["nda_note"] = ("NDA not signed, but no gas composition "
                                           "table found: inserted in full")
                        cut_pdfs.append(pdf)
                    else:
                        path, kept, total = res
                        if kept < total:
                            st2["nda_note"] = (f"NDA not signed: {total - kept} page(s) "
                                               "after the gas composition left out")
                        cut_pdfs.append(path)
                pdfs = cut_pdfs
            st2.update(paths=paths, pdfs=pdfs, conv_note="; ".join(notes))
            if len(pdfs) > 1:
                st2["path"] = merge_pdfs(pdfs, os.path.join(
                    tmp, "item_" + _re.sub(r"[^0-9A-Za-z]+", "_", code) + ".pdf"))
            elif pdfs:
                st2["path"] = pdfs[0]
            out[code] = st2
        return out

    def _commercial_values(self):
        lead = re.search(r"\d+", self.lead_time.get() or "")
        price = _num(self.final_price.get())
        code = self.incoterm.get()
        return dict(price=f"{price:,.0f}" if price > 0 else "",
                    delivery=lead.group(0) if lead else "",
                    incoterm=code, insurance_note=incoterm_note(code))

    def _preflight(self):
        """What a customer would notice as unfinished."""
        gaps = []
        missing = [st.get("title", c) for c, st in self.doc_status.items()
                   if st.get("state") == "missing" and st.get("src") == "Library"]
        if missing:
            gaps.append(f"{len(missing)} ticked document(s) not found in the "
                        f"library: {', '.join(missing[:4])}"
                        + (" …" if len(missing) > 4 else ""))
        if not self.client_name.get().strip():
            gaps.append("the customer name is empty")
        if not self.opp_id.get().strip():
            gaps.append("the proposal reference (OPP ID) is empty")
        if not self.config_code.get().strip():
            gaps.append("the compressor configuration is incomplete")
        n = self._scope_open_count()
        if n:
            gaps.append(f"{n} scope line(s) still carry a suggested supplier "
                        "nobody confirmed")
        return gaps

    def _generate_pdf(self):
        """Questions are asked first, on screen; the conversion and assembly then
        run in a worker thread so the window stays responsive (Office can take
        minutes on a large package). Progress shows in the status line."""
        if getattr(self, "_gen_busy", False):
            self._status.set("A proposal is already being generated…")
            return
        if not (HAS_REPORTLAB and HAS_PYPDF):
            missing = ("reportlab" if not HAS_REPORTLAB else "") + \
                      (" pypdf" if not HAS_PYPDF else "")
            messagebox.showerror(
                "Missing libraries",
                "The PDF assembly needs two libraries that are not installed:\n\n"
                f"    pip install {missing.strip()}\n\n"
                "Everything else in the tool keeps working without them.")
            return
        gaps = self._preflight()
        if gaps and not messagebox.askyesno(
                "Before generating",
                "This will show in the proposal:\n\n"
                + "\n".join(f"   •  {g}" for g in gaps)
                + "\n\nGenerate anyway?"):
            if self._scope_open_count():
                self._show_page(3)
                self._scope_only_open = True
                self._scope_build()
            self._status.set("Generation cancelled.")
            return
        self._scan_docs(silent=True)
        base = ("Technical_Proposal_"
                + (self.opp_id.get().strip() or "offer")
                + (f"_rev{self.revision.get().strip()}" if self.revision.get().strip() else ""))
        by_section = self.pdf_output.get() == "folder"
        start_dir = self.user_root.get().strip() or None
        skeleton = self._skeleton_path() if self.pdf_use_word.get() else ""
        if self.pdf_use_word.get() and not skeleton:
            if not messagebox.askyesno(
                    "Word master not found",
                    f"No Word file was found in:\n{_tool_dir()}\n\n"
                    "Put Technical_Proposal.docx there, or pick it with "
                    "“change”.\n\nGenerate with the built-in layout "
                    "instead?"):
                self._status.set("Generation cancelled.")
                return
        if by_section:
            parent = filedialog.askdirectory(
                title="Where to create the proposal folder", initialdir=start_dir)
            if not parent:
                return
            target = os.path.join(parent, _safe_name(base))
        else:
            target = filedialog.asksaveasfilename(
                title="Save the technical proposal", defaultextension=".pdf",
                initialfile=base + ".pdf", initialdir=start_dir,
                filetypes=[("PDF document", "*.pdf")])
            if not target:
                return

        # everything the worker needs is read here, on the UI thread
        ucp, motor = self.ucp_type.get(), self.motor_type.get()
        job = dict(
            target=target, base=base, by_section=by_section, skeleton=skeleton,
            tmp=tempfile.mkdtemp(prefix="ybtool_"),
            logo=os.path.join(_tool_dir(), "baker_hughes_logo.png"),
            options=dict(placeholders=self.pdf_placeholders.get(),
                         stamp=self.pdf_stamp.get(),
                         scope=self.pdf_scope.get() and self._lib_on("scope"),
                         commercial=self.pdf_commercial.get(),
                         skip_empty=self.pdf_skip_empty.get(),
                         frame=self.pdf_frame.get()),
            data=self._collect(), struct=self._doc_struct(output=True),
            doc_status={k: dict(v) for k, v in self.doc_status.items()},
            nda=bool(self.nda_signed.get()), comm=self._commercial_values(),
            title_of=lambda it, u=ucp, m=motor: self._doc_title(it, u, m),
            layout="word" if skeleton else "builtin")
        self._gen_start(job)

    # generation worker -------------------------------------------------------
    def _gen_start(self, job):
        import queue
        self._gen_busy = True
        self._gen_q = q = queue.Queue()
        self._gen_t0 = time.time()
        self.root.config(cursor="watch")
        self._gen_lock(True)
        self._status.set("Converting the documents…")

        def say(msg):
            q.put(("say", msg))

        def work():
            try:
                q.put(("done", self._gen_work(job, say)))
            except _AskBuiltin as e:
                q.put(("ask", e))
            except Exception as e:
                q.put(("fail", e))

        threading.Thread(target=work, daemon=True).start()
        self.root.after(80, lambda: self._gen_poll(job))

    def _gen_poll(self, job):
        try:
            while True:
                kind, val = self._gen_q.get_nowait()
                if kind == "say":
                    secs = int(time.time() - self._gen_t0)
                    self._status.set(f"{val}      {secs} s")
                    continue
                self._gen_finish(job, kind, val)
                return
        except Exception:
            pass
        try:
            self.root.after(120, lambda: self._gen_poll(job))
        except tk.TclError:
            pass

    def _gen_finish(self, job, kind, val):
        if kind == "ask":
            # the Word master failed: never change layout without asking
            if messagebox.askyesno(
                    "Word master",
                    f"{val.err}\n\nThe filled Word file is here:\n{val.filled}\n\n"
                    "Generate with the built-in layout instead?"):
                job.update(layout="builtin", how="built-in layout (Word master not used)")
                self._gen_start(job)
                return
            self._status.set("Generation stopped.")
        elif kind == "fail":
            self._status.set("Generation failed.")
            messagebox.showerror("Proposal generation", str(val))
        else:
            pages, report, how = val
            self._remember_contacts()
            self._status.set(f"Proposal generated — {pages} pages.")
            self._pdf_hint.set(f"last run: {pages} pages · {os.path.basename(job['target'])}")
            self._pdf_report(job["target"], pages, report, how)
        self._gen_busy = False
        self.root.config(cursor="")
        self._gen_lock(False)
        shutil.rmtree(job["tmp"], ignore_errors=True)

    def _gen_lock(self, on):
        """Grey the generate buttons while a run is in progress."""
        for b in getattr(self, "_gen_buttons", []):
            try:
                b.set_enabled(not on) if hasattr(b, "set_enabled") else \
                    b.config(state=tk.DISABLED if on else tk.NORMAL)
            except (tk.TclError, AttributeError):
                pass

    def _gen_wait(self, timeout=900):
        """Pump events until the current run is over (scripts and tests)."""
        t0 = time.time()
        while getattr(self, "_gen_busy", False) and time.time() - t0 < timeout:
            self.root.update()
            time.sleep(0.03)

    def _gen_work(self, job, say):
        """Runs in the worker thread: no widget, no Tk variable is touched."""
        tmp, target, by_section = job["tmp"], job["target"], job["by_section"]
        status = self._prepare_docs(tmp, job["doc_status"], job["nda"],
                                    job["comm"], say)
        data, struct, options = job["data"], job["struct"], job["options"]
        full = os.path.join(tmp, "full.pdf") if by_section else target
        builder, how = None, job.get("how", "")
        skeleton = job["skeleton"]
        if skeleton and job["layout"] == "word":
            say(f"Filling {os.path.basename(skeleton)}…")
            filled = (os.path.join(target, _safe_name(job["base"]) + ".docx")
                      if by_section else os.path.splitext(target)[0] + ".docx")
            if by_section:
                os.makedirs(target, exist_ok=True)
            try:
                builder = WordProposal(skeleton, data, struct, status,
                                       options=options, logo=job["logo"],
                                       tmp_dir=tmp, title_of=job["title_of"])
                pages = builder.build(full, keep_docx=filled)
                how = (f"{os.path.basename(skeleton)} filled · PDF by "
                       f"{builder.engine} · Word copy: {os.path.basename(filled)}")
            except Exception as e:
                raise _AskBuiltin(e, filled)
        if builder is None:
            say("Assembling the technical proposal…")
            if by_section:
                os.makedirs(target, exist_ok=True)
            builder = ProposalPDF(data, struct, status, options=options,
                                  logo=job["logo"])
            pages = builder.build(full)
            how = how or "built-in layout"
        say("Marking every page “Baker Hughes Confidential”…")
        try:
            marked = stamp_confidential(full)
            if marked:
                how += f" · “Baker Hughes Confidential” added on {marked} page(s)"
        except Exception as e:
            how += f" · confidential mark not added ({e})"
        builder.report[:0] = self._prep_report
        if self._commercial_copy:
            dst = (os.path.join(target, "Section 11_Commercial.docx")
                   if by_section else
                   os.path.splitext(target)[0] + "_Commercial.docx")
            try:
                shutil.copy2(self._commercial_copy, dst)
            except OSError:
                pass
        if by_section:
            say("Splitting into one file per section…")
            files = split_pdf_by_section(full, builder.section_starts, target)
            how += f" · {len(files)} files"
        return pages, builder.report, how

    def _pdf_report(self, out, pages, report, how=""):
        win = tk.Toplevel(self.root)
        win.title("Technical proposal generated")
        win.configure(bg=BH_WHITE)
        win.geometry("760x520")
        win.transient(self.root)

        hd = tk.Frame(win, bg=BH_INK, padx=20, pady=14)
        hd.pack(fill=tk.X)
        _lbl(hd, "Technical proposal generated", fnt=("Segoe UI", 13, "bold"),
             fg=BH_WHITE, bg=BH_INK).pack(anchor=tk.W)
        miss = sum(1 for _r, _t, o in report if o.startswith("MISSING"))
        _lbl(hd, f"{pages} pages   ·   {len(report)} entries   ·   "
                 f"{miss} document(s) still missing",
             fnt=FONT_SM, fg=BH_ON_DARK, bg=BH_INK).pack(anchor=tk.W)
        if how:
            _lbl(hd, how, fnt=FONT_SM, fg=BH_ON_DARK, bg=BH_INK,
                 wraplength=700, justify=tk.LEFT).pack(anchor=tk.W, pady=(2, 0))

        body = tk.Frame(win, bg=BH_WHITE)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)
        cv = tk.Canvas(body, bg=BH_WHITE, highlightthickness=0)
        sb = ttk.Scrollbar(body, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y); cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner = tk.Frame(cv, bg=BH_WHITE)
        wid = cv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(wid, width=e.width))

        for i, (ref, title, outcome) in enumerate(report):
            bg = BH_WHITE if i % 2 == 0 else BH_LIGHT
            row = tk.Frame(inner, bg=bg); row.pack(fill=tk.X)
            _lbl(row, ref, fnt=FONT_CS, fg=BH_GRAY, bg=bg, width=8,
                 anchor=tk.W).pack(side=tk.LEFT, pady=3)
            _lbl(row, title[:46], fnt=FONT_N, bg=bg, width=46,
                 anchor=tk.W).pack(side=tk.LEFT)
            col = (BH_RED if outcome.startswith(("MISSING", "NOT INSERTED")) else
                   BH_TOOL if "generated" in outcome or "placeholder" in outcome
                   else BH_OK)
            _lbl(row, outcome, fnt=FONT_SM, fg=col, bg=bg,
                 anchor=tk.W).pack(side=tk.LEFT, fill=tk.X, expand=True)

        ft = tk.Frame(win, bg=BH_LIGHT, pady=10, padx=16)
        ft.pack(fill=tk.X)
        if os.path.isdir(out):
            _btn(ft, "Open the folder", lambda: self._open_path(out)).pack(side=tk.LEFT)
        else:
            _btn(ft, "Open the PDF", lambda: self._open_path(out)).pack(side=tk.LEFT)
            _btn(ft, "Open the folder",
                 lambda: self._open_path(os.path.dirname(out)),
                 bg=BH_GRAY).pack(side=tk.LEFT, padx=8)
        _btn(ft, "Close", win.destroy, bg=BH_GRAY).pack(side=tk.RIGHT)

    # ── document helpers ──────────────────────────────────────────────────────
    def _layout_key(self):
        return "FIRM" if self.offer_category.get() == "FIRM" else "DTI"

    def _doc_struct(self, output=False):
        """Proposal structure in the order set in the document manager.

        Sections and sub-sections keep their identity (num, code) wherever
        they are moved; only the numbers shown and printed follow the new
        position, so the table of contents always reads 1, 2, 3…

        A line whose library documents are all unticked in “Documents to
        provide”, and that has no text of its own in the Word master, is
        out of the offer: shown at the end of its section without a number,
        and left out when output is True."""
        key = self._layout_key()
        base = DOC_STRUCT_FIRM if key == "FIRM" else DOC_STRUCT_DTI
        lay = self.doc_layout.get(key) or {}
        ticks = tuple(sorted(i for i in self.lib_vars if self._lib_on(i)))
        sig = (key, json.dumps(lay, sort_keys=True), ticks, output,
               getattr(self, "_wmap", (None,))[0])
        cache = getattr(self, "_struct_cache2", {})
        if sig in cache:
            return cache[sig]

        def ordered(seq, wanted, ident):
            by = {ident(x): x for x in seq}
            head = [by[k] for k in wanted if k in by]
            seen = set(wanted)
            return head + [x for x in seq if ident(x) not in seen]

        removed = set(lay.get("removed") or [])
        gone = set(lay.get("removed_secs") or [])
        added = lay.get("custom") or {}
        extra = [_S(e["num"], e["title"], "Section_" + _safe_name(e["title"]),
                    [dict(_D(f"X{e['num']}", e["title"], SRC_USER,
                             tuple(_title_tokens(e["title"])), ref=str(e["num"])),
                          custom=True)])
                 for e in lay.get("extra_secs") or []]
        body = [s for s in list(base) + extra
                if s["num"] != 0 and s["num"] not in gone]
        numbers = (list(range(1, len(body) + 1)) if gone or extra
                   else [s["num"] for s in body])
        out = [dict(s, disp=0) for s in base if s["num"] == 0]
        for pos, sec in enumerate(ordered(body, lay.get("secs") or [],
                                          lambda s: s["num"])):
            disp, num = numbers[pos], str(sec["num"])
            minors = [int(it["ref"].split(".")[1]) for it in sec["items"]
                      if _re.fullmatch(rf"{num}\.\d+", it["ref"])]
            nxt = min(minors) if minors else 1
            pool = [it for it in sec["items"] if it["code"] not in removed] + [
                dict(_D(c["code"], c["title"], SRC_USER,
                        tuple(_title_tokens(c["title"])), ref=f"{num}.999"),
                     custom=True)
                for c in added.get(num) or []]
            pool = ordered(pool, (lay.get("items") or {}).get(num) or [],
                           lambda i: i["code"])
            live = [it for it in pool if not self._line_off(it, sec)]
            parked = [it for it in pool if it not in live]
            items = []
            for it in live:
                ref = it["ref"]
                if ref == num:
                    ref = str(disp)
                elif _re.fullmatch(rf"{num}\.\d+", ref):
                    ref = f"{disp}.{nxt}"; nxt += 1
                kids = ordered([x for x in it["sub"] if x["code"] not in removed],
                               (lay.get("subs") or {}).get(it["code"]) or [],
                               lambda i: i["code"])
                on = [s for s in kids if not self._line_off(s, sec)]
                subs = [dict(s, ref=f"{ref}.{j}")
                        if _re.fullmatch(r"[\d.]+", s["ref"]) else dict(s)
                        for j, s in enumerate(on, 1)]
                if not output:
                    subs += [dict(s, ref="", off=True) for s in kids if s not in on]
                items.append(dict(it, ref=ref, sub=subs))
            if not output:
                items += [dict(it, ref="", off=True, sub=[]) for it in parked]
            if output and not items:
                continue
            out.append(dict(sec, items=items, disp=disp))
        cache[sig] = out
        self._struct_cache2 = cache
        return out

    def _line_off(self, it, sec):
        db = it.get("db") or []
        if not db or any(self._lib_on(i) for i in db):
            return False
        return not self._word_written(it, sec)

    def _word_map(self):
        """Titles of the Word master and whether text sits under them."""
        path = self._skeleton_path() if hasattr(self, "skeleton_file") else ""
        try:
            key = (path, os.path.getmtime(path))
        except OSError:
            return []
        if getattr(self, "_wmap", (None,))[0] == key:
            return self._wmap[1]
        out = []
        try:
            wp = WordProposal(path, {}, [], {}, tmp_dir=tempfile.gettempdir())
            els = wp._normalise(wp._load())
            wp._parse(els)
            for n in wp.sec_order:
                ws = wp.secs[n]
                out.append(dict(
                    title=ws["title"], alt=ws["alt"],
                    written=wp._written(ws["head"], ws["h1_at"]),
                    subs=[(s["title"], wp._written(s["els"], s["at"]))
                          for s in ws["subs"]]))
        except Exception:
            out = []
        self._wmap = (key, out)
        return out

    def _word_written(self, it, sec):
        best, hit = 0.0, False
        title = self._doc_title(it)
        top = it["ref"] in (str(sec["num"]), str(sec.get("disp", sec["num"])))
        for ws in self._word_map():
            near = max(_title_sim(sec["title"], ws["title"]),
                       _title_sim(sec["title"], ws["alt"])) >= 0.5
            if top:
                if near:
                    return ws["written"]
                continue
            for t2, written in ws["subs"]:
                s = max(_title_sim(it["title"], t2), _title_sim(title, t2)) \
                    + (0.05 if near else 0)
                if s > best:
                    best, hit = s, written
        return best >= 0.5 and hit

    def _reordered(self, note="Order updated — numbering and table of "
                              "contents follow."):
        self._struct_cache = (None, None)
        self._struct_cache2 = {}
        self._scan_docs(silent=True, rewalk=False)
        if hasattr(self, "_status"):
            self._status.set(note)

    def _lay(self):
        return self.doc_layout.setdefault(self._layout_key(), {})

    def _ask_title(self, heading, initial=""):
        text = simpledialog.askstring(heading, "Title printed in the proposal:",
                                      initialvalue=initial, parent=self.root)
        return (text or "").strip()

    def _add_line(self, sec_num):
        title = self._ask_title("New line")
        if not title:
            return
        lay = self._lay()
        lay["seq"] = lay.get("seq", 0) + 1
        lay.setdefault("custom", {}).setdefault(str(sec_num), []).append(
            dict(code=f"U{lay['seq']}", title=title))
        self._reordered(f"Line added: {title}")

    def _custom_entry(self, code):
        lay = self._lay()
        for rows in (lay.get("custom") or {}).values():
            for c in rows:
                if c["code"] == code:
                    return rows, c
        return None, None

    def _rename_line(self, code):
        _rows, c = self._custom_entry(code)
        if c:
            title = self._ask_title("Rename the line", c["title"])
            if title:
                c["title"] = title
                self._reordered(f"Line renamed: {title}")

    def _remove_line(self, code):
        rows, c = self._custom_entry(code)
        if c:
            rows.remove(c)
        else:
            self._lay().setdefault("removed", []).append(code)
        self.doc_attach.pop(self._att_key(code), None)
        self._reordered("Line removed — “Restore default” brings the "
                        "standard lines back.")

    def _add_section(self):
        title = self._ask_title("New section")
        if not title:
            return
        lay = self._lay()
        lay["seq"] = lay.get("seq", 0) + 1
        lay.setdefault("extra_secs", []).append(
            dict(num=100 + lay["seq"], title=title))
        self._reordered(f"Section added: {title}")

    def _remove_section(self, num):
        sec = next((s for s in self._doc_struct() if s["num"] == num), None)
        if not sec or not messagebox.askyesno(
                "Remove the section",
                f"Take “{sec['title']}” out of this proposal?"):
            return
        lay = self._lay()
        mine = [e for e in lay.get("extra_secs") or [] if e["num"] == num]
        if mine:
            lay["extra_secs"].remove(mine[0])
            (lay.get("custom") or {}).pop(str(num), None)
        else:
            lay.setdefault("removed_secs", []).append(num)
        self._reordered("Section removed — “Restore default” brings the "
                        "standard sections back.")

    def _move_section(self, num, step):
        secs = [s["num"] for s in self._doc_struct() if s["num"] != 0]
        i = secs.index(num)
        j = i + step
        if not 0 <= j < len(secs):
            return
        secs[i], secs[j] = secs[j], secs[i]
        self.doc_layout.setdefault(self._layout_key(), {})["secs"] = secs
        self._reordered()

    def _siblings(self, sec, code):
        """(codes at the same level, layout bucket name, bucket key)."""
        codes = [it["code"] for it in sec["items"]]
        if code in codes:
            return codes, "items", str(sec["num"])
        for it in sec["items"]:
            subs = [s["code"] for s in it["sub"]]
            if code in subs:
                return subs, "subs", it["code"]
        return [], "", ""

    def _move_item(self, sec_num, code, step):
        sec = next((s for s in self._doc_struct() if s["num"] == sec_num), None)
        if not sec:
            return
        codes, bucket, key = self._siblings(sec, code)
        if code not in codes:
            return
        i = codes.index(code)
        j = i + step
        if not 0 <= j < len(codes):
            return
        codes[i], codes[j] = codes[j], codes[i]
        lay = self.doc_layout.setdefault(self._layout_key(), {})
        lay.setdefault(bucket, {})[key] = codes
        self._reordered()

    def _reset_order(self):
        lay = self.doc_layout.get(self._layout_key())
        if not lay:
            return
        if (lay.get("custom") or lay.get("extra_secs")) and not messagebox.askyesno(
                "Restore default",
                "The lines and sections you added will be dropped. Continue?"):
            return
        self.doc_layout.pop(self._layout_key(), None)
        self._reordered("Standard structure restored.")

    def _dyn_keys(self, dyn):
        if dyn == "cooling":
            return [self.cooling_system.get()]
        if dyn == "casing":
            cs = _casing_digits(self.casing_size.get())
            return [f"icl{cs}", cs] if cs else []
        if dyn == "plc":
            u, out = self.ucp_type.get(), []
            if "Siemens" in u:  out += ["siemens"]
            if "Allen" in u:    out += ["rockwell", "allen bradley"]
            if "Common" in u:   out += ["common"]
            if "Separate" in u: out += ["separate"]
            return out
        if dyn == "config":
            return [self.config_code.get()]
        if dyn == "motor":
            return [self.motor_type.get()]
        return []

    def _doc_title(self, it, ucp=None, motor=None):
        """Title as printed in the proposal, adapted to the configuration."""
        if it.get("tdyn") == "plc_brand":
            u = self.ucp_type.get() if ucp is None else ucp
            brand = ("Siemens" if "Siemens" in u else
                     "Allen Bradley" if "Allen" in u else "PLC")
            return f"{brand} – PLC Configuration"
        if it.get("tdyn") == "special_tool":
            return ("Standard Special Tool – "
                    f"{self.motor_type.get() if motor is None else motor}")
        return it["title"]

    def _doc_var(self, code):
        if code not in self.doc_use_db:
            self.doc_use_db[code] = tk.BooleanVar(value=True)
        return self.doc_use_db[code]

    def _set_filter(self, key, render=True):
        self.doc_filter.set(key)
        for k, b in getattr(self, "_filter_btns", {}).items():
            on = (k == key)
            b.config(bg=BH_GREEN if on else BH_LIGHT,
                     fg=BH_WHITE if on else BH_GRAY)
        if render:
            self._apply_visibility()

    def _collapse_all(self, collapse):
        self._doc_collapsed = ({s["num"] for s in self._doc_struct()}
                               if collapse else set())
        self._apply_visibility()

    def _toggle_section(self, num):
        if num in self._doc_collapsed:
            self._doc_collapsed.discard(num)
        else:
            self._doc_collapsed.add(num)
        self._apply_visibility()

    def _pick_user_root(self):
        d = filedialog.askdirectory(
            title="Project folder (the opportunity folder is enough)")
        if d:
            self._project_pick = d
            root = proposal_root(d, self.offer_category.get())
            self.user_root.set(root)
            self._save_settings()
            self._scan_docs()
            if root != d:
                self._status.set(f"Proposal folder: {os.path.relpath(root, d)}")

    def _pick_db_root(self):
        d = filedialog.askdirectory(title="Select the standard document database")
        if d:
            self.db_root.set(d)
            self._save_settings()
            self._scan_docs()

    def _settings_file(self):
        return os.path.join(_tool_dir(), SETTINGS_FILE)

    def _load_settings(self):
        self._last_contacts = {}
        try:
            with open(self._settings_file(), encoding="utf-8") as fh:
                cfg = json.load(fh)
        except Exception:
            cfg = {}
        if cfg.get("user_root"): self.user_root.set(cfg["user_root"])
        if cfg.get("project_pick"): self._project_pick = cfg["project_pick"]
        if cfg.get("db_root") and os.path.isdir(cfg["db_root"]):
            self.db_root.set(cfg["db_root"])
        if "use_db" in cfg:      self.use_db.set(bool(cfg["use_db"]))
        if cfg.get("contacts_file"): self.contacts_file.set(cfg["contacts_file"])
        self._last_contacts = cfg.get("last_contacts") or {}
        self.lib_state = {k: dict(v) for k, v in (cfg.get("lib_state") or {}).items()}

    def _save_settings(self):
        try:
            with open(self._settings_file(), "w", encoding="utf-8") as fh:
                json.dump(dict(user_root=self.user_root.get(),
                               project_pick=self._project_pick,
                               db_root=self.db_root.get(),
                               use_db=self.use_db.get(),
                               contacts_file=self.contacts_file.get(),
                               last_contacts=getattr(self, "_last_contacts", {}),
                               lib_state=self._lib_state_now()), fh, indent=2)
        except Exception:
            pass

    def _lib_state_now(self):
        state = {k: dict(v) for k, v in self.lib_state.items()}
        if hasattr(self, "lib_vars"):
            cat = self.offer_category.get()
            mine = {d["id"]: self._lib_on(d["id"]) for d in DB_LIBRARY
                    if self._lib_on(d["id"]) != d["on"][cat]}
            if mine:
                state[cat] = mine
            else:
                state.pop(cat, None)
        return state

    def _remember_contacts(self):
        self._last_contacts = {r: getattr(self, v[0]).get().strip()
                               for r, v in self.ROLE_VARS.items()
                               if getattr(self, v[0]).get().strip()}
        self._save_settings()

    def _init_db(self):
        root = self.db_root.get().strip()
        if not root:
            messagebox.showwarning("Database", "Set a database path first."); return
        try:
            os.makedirs(root, exist_ok=True)
            have = lib_folders(root)
            for doc in DB_LIBRARY:
                if doc["id"] in have or doc["rule"] == "tool":
                    continue
                os.makedirs(os.path.join(root, _safe_name(doc["name"])),
                            exist_ok=True)
            self._save_settings(); self._scan_docs()
            messagebox.showinfo("Database ready",
                                f"Standard database initialised in:\n{root}")
        except Exception as e:
            messagebox.showerror("Database error", str(e))

    def _create_structure(self):
        root = self.user_root.get().strip()
        if not root:
            self._pick_user_root()
            root = self.user_root.get().strip()
            if not root:
                return
        cat = self.offer_category.get()
        try:
            os.makedirs(root, exist_ok=True)
            created = 0
            for sec in self._doc_struct():
                d = os.path.join(root, sec["folder"])
                if not os.path.isdir(d):
                    os.makedirs(d, exist_ok=True); created += 1
                lines = [f"SECTION {sec['disp']} — {sec['title']}", "=" * 74, ""]
                for it, depth in _leaves(sec):
                    tag = {SRC_DB: "standard document — local database",
                           SRC_USER: "TO BE PROVIDED IN THIS FOLDER",
                           SRC_TOOL: f"generated by {APP_NAME}"}[it["src"]]
                    pad = "    " * depth
                    lines.append(f"  {pad}{it['ref']:<8} "
                                 f"{self._doc_title(it):<{56 - 4*depth}} {tag}")
                with open(os.path.join(d, "_expected_documents.txt"),
                          "w", encoding="utf-8") as fh:
                    fh.write("\n".join(lines) + "\n")
            self._scan_docs()
            messagebox.showinfo(
                "Folder structure",
                f"{cat} structure ready in:\n{root}\n\n"
                f"{created} folder(s) created · "
                f"{len(self._doc_struct())} section(s) checked.")
        except Exception as e:
            messagebox.showerror("Folder structure", str(e))

    def _refresh_docs(self, *_):
        """Typing in a configuration field fires this on every keystroke:
        one silent pass, on the file lists already read."""
        if getattr(self, "_refresh_job", None):
            return

        def run():
            self._refresh_job = None
            self._scan_docs(silent=True, rewalk=False)
        self._refresh_job = self.root.after(150, run)

    def _section_dir(self, sec):
        """The section folder of the project: by its number, whatever
        follows it in the folder name."""
        root = self.user_root.get().strip()
        if not root:
            return ""
        want = os.path.join(root, sec["folder"])
        if os.path.isdir(want):
            return want
        for n in _safe_listdir(root):
            m = SECTION_DIR.match(n)
            if m and int(m.group(1)) == sec["num"] and \
                    os.path.isdir(os.path.join(root, n)):
                return os.path.join(root, n)
        return want

    def _commercial_doc(self, secdir):
        """The Word commercial section kept in Section 11_Commercial."""
        words = [os.path.join(dp, f) for dp, _d, fs in os.walk(secdir or "")
                 for f in fs if f.lower().endswith((".docx", ".docm"))
                 and not f.startswith("~$")]
        return _rank_files(words)[0] if words else ""

    def _lib_status(self, it, up, sec=None):
        ids = [i for i in it["db"] if self._lib_on(i)]
        if not ids:
            return dict(state="word", label="In the Word master", col=BH_OK,
                        icon="", path="", src="Word")
        if "scope" in ids:
            return dict(state="tool", label="Built by the tool", col=BH_TOOL,
                        icon="", path="", src="Tool")
        res = [self._lib_res.get(i) or {} for i in ids]
        files = [f for r in res for f in (r.get("files") or [])]
        gaps = [f"{DB_BY_ID[i]['title']}: {r.get('note') or 'not found'}"
                for i, r in zip(ids, res) if not r.get("files")]
        notes = [r["note"] for r in res
                 if r.get("files") and r.get("note") and r.get("quality") == "closest"]
        quals = {r.get("quality") for r in res if r.get("files")}
        if files:
            partial = bool(gaps) or "closest" in quals
            label = ("Partly found" if gaps else "Closest match" if "closest" in quals
                     else "Project folder" if quals == {"project"} else
                     "Chosen" if quals == {"chosen"} else "From the library")
            return dict(state="db", label=label,
                        col=BH_WARN if partial else BH_OK, icon="",
                        path=files[0], paths=files, src="Library", lib=True,
                        lib_note="; ".join(notes + gaps),
                        replace="cooling" in ids)
        if up:
            return dict(state="user", label="Provided", col=BH_GREEN, icon="",
                        path=up, src="Project")
        return dict(state="missing", label="Not in the library", col=BH_RED,
                    icon="", path="", src="Library", lib_note="; ".join(gaps))

    def _scan_bg(self, silent=True):
        """Read the folders in a worker thread (network shares can take seconds)
        and finish the scan on the UI thread; the window never freezes."""
        if not hasattr(self, "_doc_container"):
            return
        roots = (self.user_root.get().strip(), self.db_root.get().strip())
        self._scan_gen = getattr(self, "_scan_gen", 0) + 1
        gen = self._scan_gen
        if hasattr(self, "_status") and not silent:
            self._status.set("Reading the folders…")
        box = {}

        def work():
            try:
                box["r"] = (roots, _walk_docs(roots[0]), _walk_docs(roots[1]),
                            lib_folders(roots[1]) if roots[1] else {})
            except Exception as e:
                box["e"] = e

        th = threading.Thread(target=work, daemon=True)
        th.start()

        def poll():
            if gen != getattr(self, "_scan_gen", 0):
                return                      # a newer scan took over
            if th.is_alive():
                self.root.after(40, poll)
                return
            if "r" in box:
                r = box["r"]
                self._file_cache = (r[0], r[1], r[2])
                self._lib_folders = r[3]
                try:
                    self._scan_docs(silent=silent, rewalk=False)
                except tk.TclError:
                    pass
            else:
                self._scan_docs(silent=silent)
        self.root.after(40, poll)

    def _scan_docs(self, silent=False, rewalk=True):
        if not hasattr(self, "_doc_container"):
            return
        user_root = self.user_root.get().strip()
        db_root   = self.db_root.get().strip()
        cache = getattr(self, "_file_cache", None)
        if rewalk or not cache or cache[0] != (user_root, db_root):
            user_files, db_files = _walk_docs(user_root), _walk_docs(db_root)
            self._file_cache = ((user_root, db_root), user_files, db_files)
            self._lib_folders = lib_folders(db_root) if db_root else {}
        else:
            _roots, user_files, db_files = cache
        self._lib_update(user_files)
        lib_dirs = [d + os.sep for d in self._lib_folders.values()]
        kw_files = [p for p in db_files if not any(p.startswith(d) for d in lib_dirs)]

        self.doc_status.clear()
        dti = self._layout_key() == "DTI"
        for sec in self._doc_struct():
            secdir = self._section_dir(sec)
            in_sec, out_sec = [], []
            for pth in user_files:
                if secdir and os.path.abspath(pth).startswith(os.path.abspath(secdir)):
                    in_sec.append(pth)
                else:
                    out_sec.append(pth)

            for it, depth in _leaves(sec):
                base = dict(section=sec["num"], sec_title=sec["title"],
                            folder=sec["folder"], title=self._doc_title(it),
                            ref=it["ref"], origin=it["src"], depth=depth,
                            misplaced=False, nda=it.get("nda", False))
                if it["sub"]:                       # grouping row, no own status
                    base.update(state="group", label="", col=BH_GRAY, icon="",
                                path="", src="")
                    self.doc_status[it["code"]] = base
                    continue
                if it.get("off"):
                    base.update(state="off", label="Not in this offer",
                                col=BH_MUTED, icon="", path="", src="",
                                paths=[])
                    self.doc_status[it["code"]] = base
                    continue

                dyn = self._dyn_keys(it["dyn"])
                up, _s = _match_doc(in_sec, it["keys"], dyn) if it["keys"] else (None, 0)
                misplaced = False
                if not up and it["keys"]:
                    up, _s = _match_doc(out_sec, it["keys"], dyn)
                    misplaced = bool(up)
                use_db = self.use_db.get() and self._doc_var(it["code"]).get()

                if it.get("db"):
                    st = self._lib_status(it, up, sec)
                elif it["src"] == SRC_TOOL:
                    word = self._commercial_doc(secdir) \
                        if dti and it["code"] == "11" else ""
                    st = (dict(state="user", label="Project file", col=BH_GREEN,
                               icon="", path=word, src="Project", commercial=True)
                          if word else
                          dict(state="tool", label="Built by the tool",
                               col=BH_TOOL, icon="", path="", src="Tool"))
                elif it["src"] == SRC_USER:
                    st = (dict(state="user", label="Provided", col=BH_GREEN,
                               icon="", path=up, src="Project")
                          if up else
                          dict(state="need", label="Need to be provided",
                               col=BH_WARN, icon="", path="", src="Project"))
                else:
                    dbp, _d = _match_doc(kw_files, it["keys"], dyn) \
                        if use_db and it["keys"] else (None, 0)
                    if dbp:
                        st = dict(state="db", label="In the DB", col=BH_OK,
                                  icon="", path=dbp, src="Database")
                    elif up:
                        st = dict(state="user", label="Provided", col=BH_GREEN,
                                  icon="", path=up, src="Project")
                    elif self._word_written(it, sec):
                        st = dict(state="word", label="In the Word master",
                                  col=BH_OK, icon="", path="", src="Word")
                    elif use_db:
                        st = dict(state="missing", label="Not found in the DB",
                                  col=BH_RED, icon="", path="", src="Database")
                    else:
                        st = dict(state="need", label="Need to be provided",
                                  col=BH_WARN, icon="", path="", src="Project")
                picked = [p for p in self.doc_attach.get(self._att_key(it["code"]), [])
                          if os.path.exists(p)]
                if picked and it["src"] != SRC_TOOL:
                    st = dict(state="user", label="Attached", col=BH_GREEN,
                              icon="✓", path=picked[0], src="Attached",
                              paths=picked, attached=True,
                              commercial=st.get("commercial", False))
                st.setdefault("paths", [st["path"]] if st.get("path") else [])
                base.update(st)
                base["misplaced"] = (misplaced and st["state"] == "user"
                                     and not st.get("attached"))
                self.doc_status[it["code"]] = base

        self._render_docs()
        if not silent and hasattr(self, "_status"):
            self._status.set(f"Documents scanned — {len(user_files)} in the project "
                             f"folder, {len(self._lib_folders)} document folders in "
                             "the library.")

    def _keep(self, item, st):
        """Filter + search test for one document row."""
        if st.get("state") == "off" and not self._show_off:
            return False
        q = self.doc_search.get().strip().lower()
        if q and q not in f"{item['ref']} {self._doc_title(item)}".lower():
            return False
        mode = self.doc_filter.get()
        if item["sub"]:
            return mode == "all" and not q
        if mode == "todo":  return st.get("state") in ("need", "missing")
        if mode == "db":    return st.get("state") == "db"
        if mode == "ok":    return st.get("state") in ("db", "user", "tool", "word")
        return True

    def _hover(self, frame, base, hov):
        targets = []
        def collect(w):
            for c in w.winfo_children():
                if isinstance(c, (tk.Frame, tk.Label, tk.Canvas, tk.Checkbutton)):
                    targets.append(c); collect(c)
        collect(frame); targets.append(frame)
        frame.bind("<Enter>", lambda e: [t.config(bg=hov) for t in targets], add="+")
        frame.bind("<Leave>", lambda e: [t.config(bg=base) for t in targets], add="+")

    def _doc_sig(self):
        return (self.offer_category.get(),
                tuple((it["code"], it["ref"], bool(it.get("off")))
                      for sec in self._doc_struct() for it, _d in _leaves(sec)))

    def _render_docs(self):
        """Rebuild only when the structure changed; otherwise refresh in place."""
        sig = self._doc_sig()
        if sig != getattr(self, "_doc_sig_cache", None) or not self._doc_rows:
            self._build_doc_rows()
            self._doc_sig_cache = sig
        self._refresh_rows()
        self._apply_visibility()
        self._update_totals()

    # one-time construction ---------------------------------------------------
    def _build_doc_rows(self):
        for w in self._doc_container.winfo_children():
            w.destroy()
        for w in self._nav_holder.winfo_children():
            w.destroy()
        self._doc_rows, self._sec_rows, self._doc_order = {}, {}, []
        self._nav_chips = {}

        for sec in self._doc_struct():
            chip = tk.Label(self._nav_holder, text=f"S{sec['disp']}", font=FONT_XS,
                            bg=BH_BG, fg=BH_GRAY, padx=7, pady=3, cursor="hand2")
            chip.pack(side=tk.LEFT, padx=(0, 4))
            chip.bind("<Button-1>", lambda e, n=sec["num"]: self._jump_to(n))
            chip.bind("<Enter>", lambda e, c=chip: c.config(bg=BH_SEC))
            chip.bind("<Leave>", lambda e, c=chip, n=sec["num"]:
                      c.config(bg=self._nav_chips[n]["bg"]))
            Tooltip(chip, sec["title"])
            self._nav_chips[sec["num"]] = dict(w=chip, bg=BH_BG)

        for sec in self._doc_struct():
            head = tk.Frame(self._doc_container, bg=BH_INK, cursor="hand2")
            car = _lbl(head, "  ▾  ", fnt=FONT_SM, fg=BH_ACCENT, bg=BH_INK)
            car.pack(side=tk.LEFT)
            t1 = _lbl(head, f"SECTION {sec['disp']}", fnt=FONT_B, fg=BH_ACCENT,
                      bg=BH_INK, pady=7)
            t1.pack(side=tk.LEFT)
            t2 = _lbl(head, sec["title"], fnt=FONT_B, fg=BH_WHITE, bg=BH_INK)
            t2.pack(side=tk.LEFT, padx=(8, 0))
            if sec["num"] != 0:
                n0 = sec["num"]
                for sym, act, tip in (
                        ("✕", lambda n=n0: self._remove_section(n),
                         "Remove this section from the proposal"),
                        ("▼", lambda n=n0: self._move_section(n, 1),
                         "Move this section down"),
                        ("▲", lambda n=n0: self._move_section(n, -1),
                         "Move this section up"),
                        ("+ line", lambda n=n0: self._add_line(n),
                         "Add a line to this section")):
                    ar = _lbl(head, f" {sym} ", fnt=FONT_SM, fg=BH_ON_DARK2,
                              bg=BH_INK, cursor="hand2")
                    ar.pack(side=tk.RIGHT, padx=(0, 2))
                    ar.bind("<Button-1>", lambda e, a=act: a())
                    ar.bind("<Enter>", lambda e, a=ar: a.config(fg=BH_ACCENT))
                    ar.bind("<Leave>", lambda e, a=ar: a.config(fg=BH_ON_DARK2))
                    Tooltip(ar, tip)
            ct = _lbl(head, "0/0  ", fnt=FONT_BD, fg=BH_MUTED, bg=BH_INK)
            ct.pack(side=tk.RIGHT)
            bar = tk.Canvas(head, width=84, height=4, bg=BH_INK,
                            highlightthickness=0, bd=0)
            bar.pack(side=tk.RIGHT, padx=8)
            fd = _lbl(head, sec["folder"], fnt=FONT_CS, fg=BH_MUTED, bg=BH_INK)
            fd.pack(side=tk.RIGHT, padx=12)
            Tooltip(fd, "Folder expected inside the project folder")
            for w in (head, car, t1, t2, ct):
                w.bind("<Button-1>", lambda e, n=sec["num"]: self._toggle_section(n))
            self._sec_rows[sec["num"]] = dict(head=head, caret=car, count=ct,
                                              bar=bar, sec=sec)
            self._doc_order.append(("sec", sec["num"]))

            for i, (it, depth) in enumerate(_leaves(sec)):
                bg = BH_WHITE if i % 2 == 0 else BH_LIGHT
                row = tk.Frame(self._doc_container, bg=bg)
                tk.Frame(row, bg=BH_BORDER, height=1).pack(fill=tk.X)
                inner = tk.Frame(row, bg=bg, pady=5)
                inner.pack(fill=tk.X)

                box = tk.Frame(inner, bg=bg, width=26, height=22)
                box.pack(side=tk.LEFT); box.pack_propagate(False)
                if it.get("db") and not it["sub"] and \
                        "scope" not in it["db"]:
                    v = self.lib_vars[it["db"][0]]
                    cb = tk.Checkbutton(box, variable=v, bg=bg,
                                        selectcolor=BH_WHITE, bd=0,
                                        highlightthickness=0, activebackground=bg,
                                        command=lambda ids=it["db"], v=v:
                                        [self.lib_vars[i].set(v.get())
                                         for i in ids[1:]])
                    cb.pack()
                    Tooltip(cb, "In this offer — same tick as on the "
                                "“Documents to provide” page")
                elif it["src"] == SRC_DB and not it["sub"]:
                    cb = tk.Checkbutton(box, variable=self._doc_var(it["code"]),
                                        bg=bg, selectcolor=BH_WHITE, bd=0,
                                        highlightthickness=0, activebackground=bg,
                                        command=self._quick_refresh)
                    cb.pack()
                    Tooltip(cb, "Take this document from the standard database")
                else:
                    _lbl(box, " ", bg=bg).pack()

                off = bool(it.get("off"))
                ind = "" if depth == 0 else "    └  "
                _lbl(inner, f"{ind}{it['ref'] or '–'}", fnt=FONT_CS,
                     fg=BH_GRAY if depth == 0 and not off else BH_MUTED, bg=bg,
                     width=12, anchor=tk.W).pack(side=tk.LEFT)
                tl = _lbl(inner, self._doc_title(it),
                          fnt=FONT_B if it["sub"] else FONT_N,
                          fg=BH_MUTED if off else
                          BH_INK if depth == 0 else BH_GRAY, bg=bg,
                          width=46, anchor=tk.W)
                tl.pack(side=tk.LEFT)
                if it.get("note"):
                    tl.config(fg=BH_WARN)
                    Tooltip(tl, it["note"])

                pf = tk.Frame(inner, bg=bg, width=176, height=23)
                pf.pack(side=tk.LEFT); pf.pack_propagate(False)
                act = tk.Frame(inner, bg=bg)
                act.pack(side=tk.RIGHT, padx=(6, 8))
                lf = _lbl(inner, "—", fnt=FONT_SM, fg=BH_MUTED, bg=bg, anchor=tk.W)
                lf.pack(side=tk.LEFT, fill=tk.X, expand=True)

                self._hover(inner, bg, BH_SEC)
                for w in (inner, tl, lf):
                    w.bind("<Double-Button-1>",
                           lambda e, cd=it["code"]: self._open_doc(cd))
                    w.bind("<Button-3>",
                           lambda e, s=sec, i2=it: self._row_menu(e, s, i2))
                self._doc_rows[it["code"]] = dict(
                    row=row, inner=inner, pill=pf, file=lf, act=act,
                    item=it, sec=sec, bg=bg, title=tl)
                self._doc_order.append(("row", it["code"]))

    # in-place refresh --------------------------------------------------------
    def _row_sig(self, code, w):
        st = self.doc_status.get(code, {})
        it, sec = w["item"], w["sec"]
        subs = tuple(self.doc_status.get(s["code"], {}).get("state")
                     for s in it["sub"]) if it["sub"] else ()
        libs = tuple(len((self._lib_res.get(i) or {}).get("candidates") or [])
                     for i in (it.get("db") or ()))
        sib = (len(self._siblings(sec, code)[0]) > 1) if sec["num"] != 0 else False
        return (st.get("label"), st.get("col"), st.get("state"),
                tuple(st.get("paths") or ()), st.get("path"), st.get("lib_note"),
                st.get("misplaced"), self._doc_title(it), bool(it.get("off")),
                subs, libs, sib)

    def _refresh_rows(self):
        for code, w in self._doc_rows.items():
            sig = self._row_sig(code, w)
            if w.get("_sig") == sig:
                continue
            w["_sig"] = sig
            st = self.doc_status.get(code, {})
            it, bg = w["item"], w["bg"]
            for child in w["pill"].winfo_children():
                child.destroy()
            if it["sub"]:
                ok = sum(1 for s in it["sub"]
                         if self.doc_status.get(s["code"], {}).get("state")
                         in ("db", "user", "tool", "word"))
                _pill(w["pill"], f"{ok} of {len(it['sub'])} ready", BH_MUTED,
                      host_bg=bg).pack(anchor=tk.W)
            else:
                _pill(w["pill"], st.get("label", "—"), st.get("col", BH_GRAY),
                      host_bg=bg, icon=st.get("icon", "")).pack(anchor=tk.W)

            paths = st.get("paths") or ([st["path"]] if st.get("path") else [])
            name = ",  ".join(os.path.basename(p) for p in paths[:2])
            if len(paths) > 2:
                name += f"   +{len(paths) - 2}"
            name = name or ("" if it["sub"] else "—")
            if st.get("lib_note"):
                name += f"    {st['lib_note']}"
            if st.get("misplaced"):
                name += "    ⚠ outside its section folder"
            w["file"].config(text=name,
                             fg=BH_INK if st.get("path") else BH_MUTED)
            if st.get("lib_note") and not w.get("_tip"):
                w["_tip"] = Tooltip(w["file"], st["lib_note"])
            elif w.get("_tip") is not None:
                w["_tip"].text = st.get("lib_note") or ""
            w["title"].config(text=self._doc_title(it))

            for child in w["act"].winfo_children():
                child.destroy()
            if st.get("path"):
                _ghost_btn(w["act"], "Open",
                           lambda c=code: self._open_doc(c), bg=bg).pack(side=tk.RIGHT)
            if it.get("db") and not it.get("off") and any(
                    len((self._lib_res.get(i) or {}).get("candidates") or []) > 1
                    and DB_BY_ID[i]["rule"] != "all" for i in it["db"]):
                lid = next(i for i in it["db"]
                           if len((self._lib_res.get(i) or {}).get("candidates") or []) > 1)
                _ghost_btn(w["act"], "Change",
                           lambda i=lid, a=w["act"]: self._lib_menu(i, a),
                           bg=bg).pack(side=tk.RIGHT, padx=4)
            if it["src"] != SRC_TOOL and not it["sub"] and not it.get("off"):
                _ghost_btn(w["act"], "Attach",
                           lambda s=w["sec"], i2=it: self._attach_file(s, i2),
                           fg=BH_TOOL, bg=bg).pack(side=tk.RIGHT, padx=4)
            sec = w["sec"]
            if sec["num"] != 0 and not it.get("off") and \
                    len(self._siblings(sec, code)[0]) > 1:
                for sym, step in (("▼", 1), ("▲", -1)):
                    _ghost_btn(w["act"], sym,
                               lambda n=sec["num"], c=code, s=step:
                               self._move_item(n, c, s),
                               fg=BH_MUTED, bg=bg).pack(side=tk.RIGHT)
            if sec["num"] != 0 and not it.get("off"):
                rm = _ghost_btn(w["act"], "✕",
                                lambda c=code: self._remove_line(c),
                                fg=BH_MUTED, bg=bg)
                rm.pack(side=tk.RIGHT)
                Tooltip(rm, "Remove this line from the proposal")

    # visibility (collapse + filter + search) ---------------------------------
    def _jump_to(self, num):
        self._doc_collapsed.discard(num)
        self._apply_visibility()
        sr = self._sec_rows.get(num)
        if sr:
            self._scroll_to(sr["head"])

    def _scroll_to(self, widget):
        page = self.t_docmgr
        sc = getattr(page, "_scroller", None)
        if sc is None:
            return
        page.update_idletasks()
        sc.to_widget(widget, margin=20)

    def _apply_visibility(self):
        shown = 0
        for kind, key in self._doc_order:
            (self._sec_rows[key]["head"] if kind == "sec"
             else self._doc_rows[key]["row"]).pack_forget()

        for sec in self._doc_struct():
            num = sec["num"]
            sr = self._sec_rows.get(num)
            if not sr:
                continue
            collapsed = num in self._doc_collapsed
            visible = [it["code"] for it, _d in _leaves(sec)
                       if self._keep(it, self.doc_status.get(it["code"], {}))]
            chip = self._nav_chips.get(num)
            if not visible and (self.doc_filter.get() != "all"
                                or self.doc_search.get().strip()):
                if chip:
                    chip["w"].pack_forget()
                continue
            if chip:
                chip["w"].pack(side=tk.LEFT, padx=(0, 4))
            shown += len(visible)
            sr["caret"].config(text="  ▸  " if collapsed else "  ▾  ")
            sr["head"].pack(fill=tk.X, pady=(12, 0))
            if collapsed:
                continue
            q = self.doc_search.get().strip().lower()
            for code in visible:
                w = self._doc_rows[code]
                w["row"].pack(fill=tk.X)
                hit = q and q in self._doc_title(w["item"]).lower()
                w["title"].config(fg=BH_GREEN if hit else
                                  (BH_INK if not w["item"]["sub"] else BH_INK))

        self._doc_empty.pack_forget()
        if not shown:
            self._doc_empty.pack(fill=tk.X, pady=10)

    def _toggle_off(self):
        self._show_off = not self._show_off
        self._apply_visibility()
        self._update_totals()

    def _update_totals(self):
        if hasattr(self, "_off_toggle"):
            n = sum(1 for st in self.doc_status.values() if st.get("state") == "off")
            self._off_toggle.config(
                text=("" if not n else
                      f"Hide the {n} lines not in this offer" if self._show_off else
                      f"Show the {n} lines not in this offer"))
        ready = total = 0
        for sec in self._doc_struct():
            leaf = [it for it, _d in _leaves(sec)
                    if not it["sub"] and not it.get("off")]
            done = sum(1 for it in leaf
                       if self.doc_status.get(it["code"], {}).get("state")
                       in ("db", "user", "tool", "word"))
            total += len(leaf); ready += done
            sr = self._sec_rows.get(sec["num"])
            if not sr:
                continue
            sr["count"].config(text=f"{done}/{len(leaf)}  ",
                               fg=BH_OK if done == len(leaf) else BH_WARN)
            bar = sr["bar"]; bar.delete("all")
            bar.create_rectangle(0, 0, 84, 4, fill=BH_LINE_D, outline="")
            if leaf:
                bar.create_rectangle(0, 0, 84*done/len(leaf), 4,
                                     fill=BH_OK if done == len(leaf) else BH_G1,
                                     outline="")

        counts = dict(all=0, todo=0, db=0, ok=0)
        for st in self.doc_status.values():
            stt = st.get("state")
            if stt in ("group", "off"):
                continue
            counts["all"] += 1
            if stt in ("need", "missing"): counts["todo"] += 1
            if stt == "db":                counts["db"] += 1
            if stt in ("db", "user", "tool", "word"): counts["ok"] += 1
        for key, btn in getattr(self, "_filter_btns", {}).items():
            btn.config(text=f"  {self._filter_txt[key]}  {counts.get(key, 0)}  ")

        for sec in self._doc_struct():
            chip = self._nav_chips.get(sec["num"])
            if not chip:
                continue
            leaf = [it for it, _d in _leaves(sec)
                    if not it["sub"] and not it.get("off")]
            done = sum(1 for it in leaf
                       if self.doc_status.get(it["code"], {}).get("state")
                       in ("db", "user", "tool", "word"))
            full = leaf and done == len(leaf)
            chip["bg"] = BH_OK_BG if full else BH_WARN_BG
            chip["w"].config(bg=chip["bg"],
                             fg=BH_OK if full else BH_WARN,
                             text=f"S{sec['disp']}  {done}/{len(leaf)}")

        pct = int(ready / total * 100) if total else 0
        self._doc_progress["value"] = pct
        if hasattr(self, "_rail_ready"):
            self._rail_ready.set(f"{ready} / {total} documents")
            self._rail_bar["value"] = pct
        if hasattr(self, "_pdf_hint"):
            self._pdf_hint.set(f"{ready} of {total} documents ready")
        miss = sorted(c for c, st in self.doc_status.items()
                      if st.get("state") in ("need", "missing"))
        self._doc_summary.set(
            f"{ready} of {total} documents available  ({pct} %)"
            + (f"      still to provide:  {',  '.join(miss[:14])}"
               + (" …" if len(miss) > 14 else "")
               if miss else "      the package is complete"))

    # row interactions --------------------------------------------------------
    def _quick_refresh(self):
        """Checkbox toggled: recompute from the cached file lists, no re-walk."""
        self._refresh_docs()

    def _open_doc(self, code):
        st = self.doc_status.get(code, {})
        if st.get("path"):
            self._open_path(st["path"])

    def _row_menu(self, event, sec, item):
        st = self.doc_status.get(item["code"], {})
        m = tk.Menu(self.root, tearoff=0, bg=BH_WHITE, fg=BH_INK,
                    activebackground=BH_GREEN, activeforeground=BH_WHITE,
                    font=FONT_N, bd=0)
        if st.get("path"):
            m.add_command(label="Open the document",
                          command=lambda: self._open_path(st["path"]))
            m.add_command(label="Open the containing folder",
                          command=lambda: self._open_path(os.path.dirname(st["path"])))
            m.add_command(label="Copy the full path",
                          command=lambda: (self.root.clipboard_clear(),
                                           self.root.clipboard_append(st["path"])))
            m.add_separator()
        if item["src"] != SRC_TOOL and not item["sub"]:
            m.add_command(label="Attach files here…",
                          command=lambda: self._attach_file(sec, item))
            if self._att_key(item["code"]) in self.doc_attach:
                m.add_command(label="Remove the attached files",
                              command=lambda: self._detach_file(item))
        if item.get("custom"):
            m.add_command(label="Rename…",
                          command=lambda: self._rename_line(item["code"]))
        for lid in item.get("db") or []:
            if len((self._lib_res.get(lid) or {}).get("candidates") or []) > 1 \
                    and DB_BY_ID[lid]["rule"] != "all":
                m.add_command(label=f"Use another {DB_BY_ID[lid]['name']} file…",
                              command=lambda i=lid: self._lib_menu(i, event=event))
        m.add_command(label="Open the section folder",
                      command=lambda: self._open_path(self._section_dir(sec)))
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def _attach_file(self, sec, item):
        """Pick one or several files anywhere on the computer or the network.
        They are used from where they are; the row shows their names."""
        picked = filedialog.askopenfilenames(
            title=f"Files for {item['ref']} — {self._doc_title(item)}",
            filetypes=[("Documents", "*.pdf *.xlsx *.xlsm *.xls *.docx *.doc "
                                     "*.pptx *.png *.jpg *.jpeg *.tif *.tiff"),
                       ("All files", "*.*")])
        if not picked:
            return
        self.doc_attach[self._att_key(item["code"])] = [os.path.abspath(p) for p in picked]
        self._scan_docs(silent=True, rewalk=False)
        names = ", ".join(os.path.basename(p) for p in picked)
        self._status.set(f"Attached to {item['ref']}: {names}")

    def _att_key(self, code):
        # the same code means different documents in the two structures
        return f"{self._layout_key()}:{code}"

    def _detach_file(self, item):
        if self.doc_attach.pop(self._att_key(item["code"]), None) is not None:
            self._scan_docs(silent=True, rewalk=False)
            self._status.set(f"Attachment removed from {item['ref']}.")

    def _fill_from_db(self):
        root = self.user_root.get().strip()
        if not root:
            messagebox.showwarning("Project folder",
                                   "Select the project folder first."); return
        if not self.use_db.get():
            messagebox.showinfo("Database off",
                                "Turn on “Use standard documents from the local "
                                "database” first."); return
        if not self.auto_place.get():
            messagebox.showinfo("Copy disabled",
                                "Turn on “Copy them into the project folder” to "
                                "place the standard documents."); return
        copied = skipped = 0
        try:
            for sec in self._doc_struct():
                for it, _d in _leaves(sec):
                    st = self.doc_status.get(it["code"], {})
                    if st.get("state") != "db" or not st.get("path"):
                        continue
                    dst_dir = os.path.join(root, sec["folder"])
                    os.makedirs(dst_dir, exist_ok=True)
                    ext = os.path.splitext(st["path"])[1]
                    hint = self._doc_title(it).replace("/", "-").replace("&", "and")
                    dst = os.path.join(dst_dir, f"{it['code']}_{hint}{ext}")
                    if os.path.exists(dst):
                        skipped += 1
                    else:
                        shutil.copy2(st["path"], dst); copied += 1
            self._scan_docs()
            messagebox.showinfo("Fill from database",
                                f"{copied} standard document(s) copied into the "
                                f"project folder.\n{skipped} already in place.")
        except Exception as e:
            messagebox.showerror("Fill from database", str(e))

    def _open_path(self, path):
        path = (path or "").strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Open", "This path does not exist yet."); return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)                       # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Open", str(e))
    # ── PAGE 5: SCOPE OF SUPPLY ───────────────────────────────────────────────
    # ── TAB 5: SCOPE OF SUPPLY ────────────────────────────────────────────────
    W_ITEM, W_REMARK = 40, 24       # column widths, in characters
    PILLS = (("bh", "BH", 4), ("cust", "Customer", 9), ("opt", "Option", 7))

    def _tab_scope(self):
        p = self.t_scope
        f = self._section(p, "Scope of Supply",
                          "one supplier per line  ·  printed as shown here")

        top = tk.Frame(f, bg=BH_WHITE); top.pack(fill=tk.X, pady=(0, 10))
        self._scope_sum = tk.StringVar()
        tk.Label(top, textvariable=self._scope_sum, font=FONT_B, bg=BH_WHITE,
                 fg=BH_INK).pack(side=tk.LEFT)
        self._scope_todo = tk.Label(top, text="", font=FONT_B, bg=BH_WHITE,
                                    fg=BH_WARN)
        self._scope_todo.pack(side=tk.LEFT, padx=(18, 0))
        cb = _btn(top, "Confirm all as shown", self._scope_confirm_all,
                  bg=BH_GRAY)
        cb.pack(side=tk.RIGHT)
        Tooltip(cb, "Accept the suggested supplier on every line still "
                    "to confirm")
        self._scope_filter = tk.Label(top, text="  To confirm only  ",
                                      font=FONT_BD, bg=BH_LIGHT, fg=BH_GRAY,
                                      padx=6, pady=5, cursor="hand2")
        self._scope_filter.pack(side=tk.RIGHT, padx=8)
        self._scope_filter.bind("<Button-1>", lambda e: self._scope_toggle_filter())

        hdr = tk.Frame(f, bg=BH_GREEN); hdr.pack(fill=tk.X)
        tk.Frame(hdr, bg=BH_GREEN, width=34, height=1).pack(side=tk.LEFT)
        tk.Label(hdr, text="Item", font=FONT_B, bg=BH_GREEN, fg=BH_WHITE,
                 width=self.W_ITEM, anchor=tk.W, pady=5).pack(side=tk.LEFT)
        tk.Frame(hdr, bg=BH_GREEN, width=30, height=1).pack(side=tk.RIGHT)
        tk.Label(hdr, text="Internal note  ·  never printed", font=FONT_SM,
                 bg=BH_DARK, fg=BH_ON_DARK, width=self.W_REMARK + 4, anchor=tk.W,
                 padx=4, pady=5).pack(side=tk.RIGHT, padx=(10, 0))
        for _val, txt, w in reversed(self.PILLS):
            tk.Label(hdr, text=txt, font=FONT_SM, bg=BH_GREEN, fg=BH_WHITE,
                     width=w, padx=4).pack(side=tk.RIGHT, padx=1)
        tk.Label(hdr, text="Description  ·  printed for the customer", font=FONT_B,
                 bg=BH_GREEN, fg=BH_WHITE, anchor=tk.W).pack(side=tk.LEFT, padx=(6, 0))

        self._scope_frame = tk.Frame(f, bg=BH_WHITE)
        self._scope_frame.pack(fill=tk.X)
        self._scope_build()

        btn_row = tk.Frame(f, bg=BH_WHITE, pady=8); btn_row.pack(fill=tk.X)
        _btn(btn_row, "Add a line", self._add_scope_row).pack(side=tk.LEFT)
        _btn(btn_row, "Customer view", self._scope_preview,
             bg=BH_GRAY).pack(side=tk.LEFT, padx=(8, 0))
        _lbl(btn_row, "Item and description are what the customer reads; the "
                      "internal note stays in the tool and in your saved file. "
                      "A description left as proposed keeps following the "
                      "technical page.", fnt=FONT_SM, fg=BH_MUTED,
             wraplength=760, justify=tk.LEFT).pack(side=tk.LEFT, padx=14)

    # ── scope model ───────────────────────────────────────────────────────────
    def _scope_row(self, item):
        return dict(item=item, included=tk.BooleanVar(value=True),
                    resp=tk.StringVar(value=item["resp"]["DTI"]),
                    desc=tk.StringVar(), desc_auto="",
                    comment=tk.StringVar(), touched=False, na=False,
                    forced=None, w=None)

    def _scope_sync(self, *_):
        """Bring every line in step with the technical page: description,
        whether the line exists for this configuration, and the supplier
        where a choice made there settles it."""
        cat = self.offer_category.get()
        for row in self.scope_rows:
            it = row["item"]
            try:
                new = it["desc"](self) if callable(it["desc"]) else it["desc"]
            except Exception:
                new = ""
            new = (new or "").strip()
            if row["desc"].get() == row["desc_auto"]:   # not edited by hand
                row["desc"].set(new)
            row["desc_auto"] = new
            try:
                row["na"] = bool(it["applies"]) and not it["applies"](self)
            except Exception:
                row["na"] = False
            try:
                forced = it["auto"](self) if it["auto"] else None
            except Exception:
                forced = None
            if forced != row["forced"]:
                row["forced"] = forced
                if forced:
                    row["resp"].set(forced)
                elif not row["touched"]:
                    row["resp"].set(it["resp"].get(cat, "bh"))
        self._scope_paint()

    def _scope_open(self, row):
        """Still waiting for somebody to say who supplies it."""
        it = row.get("item")
        return bool(it and it["ask"] and row["included"].get()
                    and not row["na"] and not row["touched"]
                    and not row["forced"])

    INTERNAL_WORDS = (r"tbc", r"tbd", r"to be checked", r"to check", r"check with",
                      r"à vérifier", r"a verifier", r"à confirmer", r"a confirmer",
                      r"internal", r"interne", r"ask", r"demander", r"todo", r"xxx",
                      r"\?\?", r"voir avec", r"see with", r"not for (?:the )?customer",
                      r"margin", r"marge", r"our cost", r"notre prix")

    def _scope_suspect(self, row):
        """A description typed by hand that reads like a note to oneself
        rather than text for the customer. Proposed texts are never flagged."""
        raw = row["desc"].get().strip()
        if not raw or (row.get("item") and raw == row.get("desc_auto", "")):
            return ""
        d = raw.lower()
        for w_ in self.INTERNAL_WORDS:
            m = re.search(r"(?<!\w)" + w_ + r"(?!\w)", d)
            if m:
                return m.group(0)
        name = (row["item"]["item"] if row.get("item") else
                row["label_var"].get()).strip().lower()
        if name and d == name:
            return "same as the item"
        return ""

    def _scope_flag(self, row):
        w = row.get("w")
        if not w:
            return
        why = self._scope_suspect(row)
        try:
            w["desc"].config(highlightbackground=BH_WARN if why else BH_BORDER,
                             highlightthickness=2 if why else 1)
        except tk.TclError:
            pass

    def _scope_suspects(self):
        return [(r, self._scope_suspect(r)) for r in self.scope_rows + self.extra_scope_rows
                if r["included"].get() and not r.get("na") and self._scope_suspect(r)]

    def _scope_preview(self):
        """The table exactly as the customer will read it."""
        lines = self._scope_out()
        win = tk.Toplevel(self.root)
        win.title("Scope of supply — customer view")
        win.configure(bg=BH_WHITE)
        win.geometry("900x640")
        win.transient(self.root)
        hd = tk.Frame(win, bg=BH_INK, padx=18, pady=12); hd.pack(fill=tk.X)
        _lbl(hd, "Scope of supply — as printed in the proposal", fnt=FONT_S,
             fg=BH_WHITE, bg=BH_INK).pack(anchor=tk.W)
        sus = self._scope_suspects()
        _lbl(hd, (f"{len(lines)} lines  ·  internal notes are not shown"
                  + (f"  ·  {len(sus)} description(s) to read again" if sus else "")),
             fnt=FONT_SM, fg=BH_ON_DARK, bg=BH_INK).pack(anchor=tk.W)
        body = tk.Frame(win, bg=BH_WHITE); body.pack(fill=tk.BOTH, expand=True,
                                                     padx=14, pady=10)
        cv = tk.Canvas(body, bg=BH_WHITE, highlightthickness=0)
        sb = ttk.Scrollbar(body, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y); cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner = tk.Frame(cv, bg=BH_WHITE)
        cv.create_window((0, 0), window=inner, anchor="nw", width=840)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        has_opt = any(l["resp"] == "opt" for l in lines)
        cols = [("bh", "Baker Hughes"), ("cust", "Customer")] + \
               ([("opt", "Option")] if has_opt else [])
        h = tk.Frame(inner, bg=BH_GREEN); h.pack(fill=tk.X)
        tk.Label(h, text="Item", font=FONT_B, bg=BH_GREEN, fg=BH_WHITE,
                 anchor=tk.W, padx=8, pady=5).pack(side=tk.LEFT, fill=tk.X, expand=True)
        for _k, t in reversed(cols):
            tk.Label(h, text=t, font=FONT_SM, bg=BH_GREEN, fg=BH_WHITE,
                     width=12, pady=5).pack(side=tk.RIGHT)
        bad = {id(r) for r, _w in sus}
        rows_by = {(r["item"]["item"] if r.get("item") else r["label_var"].get().strip()): r
                   for r in self.scope_rows + self.extra_scope_rows}
        cur = None
        for l in lines:
            if l["section"] != cur:
                cur = l["section"]
                tk.Label(inner, text=cur, font=FONT_B, bg=BH_BG, fg=BH_INK,
                         anchor=tk.W, padx=8, pady=3).pack(fill=tk.X, pady=(6, 0))
            rr = tk.Frame(inner, bg=BH_WHITE); rr.pack(fill=tk.X)
            tk.Frame(inner, bg=BH_BORDER, height=1).pack(fill=tk.X)
            left = tk.Frame(rr, bg=BH_WHITE); left.pack(side=tk.LEFT, fill=tk.X,
                                                        expand=True, padx=8, pady=3)
            _lbl(left, l["item"], fnt=FONT_N).pack(anchor=tk.W)
            if l["desc"]:
                warn = id(rows_by.get(l["item"])) in bad
                _lbl(left, l["desc"] + ("     ⚠ reads like an internal note" if warn else ""),
                     fnt=FONT_SM, fg=BH_WARN if warn else BH_GRAY,
                     wraplength=560, justify=tk.LEFT).pack(anchor=tk.W)
            for k, _t in reversed(cols):
                tk.Label(rr, text="X" if l["resp"] == k else "",
                         font=(FONT_SM[0], FONT_SM[1], "bold"),
                         bg=BH_WHITE, fg=BH_INK, width=12).pack(side=tk.RIGHT)

    def _scope_open_count(self):
        return sum(1 for r in self.scope_rows if self._scope_open(r))

    def _scope_pick(self, row, val):
        if row["na"]:
            return
        row["resp"].set(val)
        row["touched"] = True
        self._scope_paint()

    def _scope_confirm_all(self):
        n = 0
        for row in self.scope_rows:
            if self._scope_open(row):
                row["touched"] = True; n += 1
        if self._scope_only_open:
            self._scope_only_open = False
            self._scope_build()
        else:
            self._scope_paint()
        self._status.set(f"{n} scope line(s) confirmed." if n
                         else "Nothing left to confirm.")

    def _scope_toggle_filter(self):
        self._scope_only_open = not self._scope_only_open
        self._scope_build()

    def _resolve(self, item):
        """Printed label of a standard line (item and description)."""
        row = next((r for r in self.scope_rows if r["item"] is item), None)
        desc = row["desc"].get().strip() if row else ""
        return item["item"] + (f"  —  {desc}" if desc else "")

    def _render_scope(self, _=None):
        self._scope_sync()

    # ── scope view ────────────────────────────────────────────────────────────
    def _scope_build(self):
        if not hasattr(self, "_scope_frame"):
            return
        for w in self._scope_frame.winfo_children():
            w.destroy()
        self._scope_secs = {}
        shown = [r for r in self.scope_rows
                 if not self._scope_only_open or self._scope_open(r)]
        current = None
        for row in self.scope_rows + self.extra_scope_rows:
            row["w"] = None
        for row in shown:
            sec = row["item"]["section"]
            if sec != current:
                current = sec
                self._scope_head(sec)
            self._scope_line_widgets(row, row["item"]["item"])
        if not self._scope_only_open:
            if self.extra_scope_rows:
                self._scope_head("OTHER")
            for row in self.extra_scope_rows:
                self._scope_line_widgets(row, None)
        elif not shown:
            box = tk.Frame(self._scope_frame, bg=BH_LIGHT, pady=22)
            box.pack(fill=tk.X, pady=(6, 0))
            _lbl(box, "Every line has a supplier.", fnt=FONT_S, fg=BH_GRAY,
                 bg=BH_LIGHT).pack()
        self._scope_paint()

    def _scope_head(self, title):
        sh = tk.Frame(self._scope_frame, bg=BH_INK)
        sh.pack(fill=tk.X, pady=(8, 0))
        tk.Label(sh, text=f"   {title}", font=FONT_B, bg=BH_INK, fg=BH_WHITE,
                 pady=5).pack(side=tk.LEFT)
        cnt = tk.Label(sh, text="", font=FONT_SM, bg=BH_INK, fg=BH_ON_DARK)
        cnt.pack(side=tk.RIGHT, padx=10)
        self._scope_secs[title] = cnt

    def _scope_line_widgets(self, row, item_text):
        """One line. item_text is None for a line added by hand, whose item
        is typed in."""
        bg = BH_WHITE
        wrap = tk.Frame(self._scope_frame, bg=bg)
        wrap.pack(fill=tk.X)
        tk.Frame(wrap, bg=BH_BORDER, height=1).pack(fill=tk.X)
        rf = tk.Frame(wrap, bg=bg)
        rf.pack(fill=tk.X)
        mark = tk.Frame(rf, bg=bg, width=4)
        mark.pack(side=tk.LEFT, fill=tk.Y)
        inc = tk.Checkbutton(rf, variable=row["included"], bg=bg,
                             selectcolor=BH_WHITE, activebackground=bg, bd=0,
                             highlightthickness=0, command=self._scope_paint)
        inc.pack(side=tk.LEFT, padx=(4, 2), pady=3)
        Tooltip(inc, "Untick to leave this line out of the proposal")
        if item_text is None:
            name = _entry(rf, row["label_var"], w=self.W_ITEM)
            name.pack(side=tk.LEFT, pady=3)
        else:
            name = tk.Label(rf, text=item_text, font=FONT_N, bg=bg, fg=BH_INK,
                            width=self.W_ITEM, anchor=tk.W, justify=tk.LEFT,
                            wraplength=int(self.W_ITEM * 6.4 * UI["scale"]))
            name.pack(side=tk.LEFT, pady=3)
            if row["item"]["hint"]:
                Tooltip(name, row["item"]["hint"])

        tail = tk.Frame(rf, bg=bg, width=30, height=1)
        tail.pack(side=tk.RIGHT)
        if item_text is None:
            def remove(r=row):
                self.extra_scope_rows = [x for x in self.extra_scope_rows
                                         if x is not r]
                self._scope_build()
            tail.pack_forget()
            _ghost_btn(rf, "✕", remove, fg=BH_MUTED, bg=bg).pack(side=tk.RIGHT)
        note = tk.Entry(rf, textvariable=row["comment"], width=self.W_REMARK + 4,
                        font=(FONT_SM[0], FONT_SM[1], "italic"), relief=tk.FLAT,
                        bd=0, bg=BH_BG, fg=BH_GRAY, insertbackground=BH_GRAY,
                        highlightthickness=1, highlightbackground=BH_BG,
                        highlightcolor=BH_GRAY)
        note.pack(side=tk.RIGHT, padx=(10, 0), pady=3, ipady=2)
        Tooltip(note, "Internal note: for you and the team. Never printed in "
                      "the proposal nor in the customer Excel.")
        pills = {}
        for val, txt, w in reversed(self.PILLS):
            pl = tk.Label(rf, text=txt, font=FONT_SM, width=w, padx=4, pady=3,
                          cursor="hand2")
            pl.pack(side=tk.RIGHT, padx=1)
            pl.bind("<Button-1>", lambda e, r=row, v=val: self._scope_pick(r, v))
            pills[val] = pl
        desc = _entry(rf, row["desc"], w=10)
        desc.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 8), pady=3)
        tip = Tooltip(desc, "")
        def printed(*_a, r=row, t=tip):
            name_ = (r["item"]["item"] if r.get("item") else r["label_var"].get().strip())
            d_ = r["desc"].get().strip()
            t.text = (f"Printed for the customer:  {name_}" + (f"  —  {d_}" if d_ else ""))
            self._scope_flag(r)
        row["desc"].trace_add("write", printed)
        row["w"] = dict(mark=mark, name=name, pills=pills, desc=desc,
                        label=item_text is not None)
        printed()

    PILL_ON = {"bh": BH_GREEN, "cust": BH_INK, "opt": BH_WARN}

    def _scope_paint(self):
        """Colours and counters, without rebuilding anything."""
        if not hasattr(self, "_scope_frame"):
            return
        per, total = {}, dict(bh=0, cust=0, opt=0)
        for row in self.scope_rows + self.extra_scope_rows:
            sec = row["item"]["section"] if row.get("item") else "OTHER"
            live = row["included"].get() and not row.get("na")
            if row.get("item") is None and not row["label_var"].get().strip():
                live = False
            resp = row["resp"].get()
            if live:
                per.setdefault(sec, dict(bh=0, cust=0, opt=0))[resp] += 1
                total[resp] += 1
            w = row.get("w")
            if not w:
                continue
            todo = self._scope_open(row)
            sig = (live, resp, todo, row.get("na"))
            if w.get("_sig") == sig:
                continue
            w["_sig"] = sig
            w["mark"].config(bg=BH_WARN if todo else BH_WHITE)
            if w["label"]:
                text = row["item"]["item"]
                if row["na"]:
                    text += "   (not in this configuration)"
                w["name"].config(text=text,
                                 fg=BH_INK if live else BH_MUTED)
            for val, pl in w["pills"].items():
                if not live:
                    pl.config(bg=BH_BG, fg=BH_GRID)
                elif val != resp:
                    pl.config(bg=BH_BG, fg=BH_MUTED)
                elif todo:                      # suggested, not confirmed
                    pl.config(bg=BH_WARN_BG, fg=BH_WARN)
                else:
                    pl.config(bg=self.PILL_ON[val], fg=BH_WHITE)
        for sec, lbl in getattr(self, "_scope_secs", {}).items():
            c = per.get(sec, dict(bh=0, cust=0, opt=0))
            bits = [f"BH {c['bh']}", f"Customer {c['cust']}"]
            if c["opt"]:
                bits.append(f"Option {c['opt']}")
            try:
                lbl.config(text="   ·   ".join(bits))
            except tk.TclError:
                pass
        n = self._scope_open_count()
        if hasattr(self, "_scope_sum"):
            self._scope_sum.set(
                f"{sum(total.values())} lines      Baker Hughes {total['bh']}"
                f"      Customer {total['cust']}      Option {total['opt']}")
            self._scope_todo.config(
                text=f"{n} to confirm" if n else "all suppliers confirmed",
                fg=BH_WARN if n else BH_OK)
            on = self._scope_only_open
            self._scope_filter.config(bg=BH_GREEN if on else BH_LIGHT,
                                      fg=BH_WHITE if on else BH_GRAY)

    def _add_scope_row(self):
        row = dict(label_var=tk.StringVar(), included=tk.BooleanVar(value=True),
                   resp=tk.StringVar(value="bh"), desc=tk.StringVar(),
                   desc_auto="", comment=tk.StringVar(), touched=True,
                   na=False, forced=None, w=None)
        row["label_var"].trace_add("write", lambda *_: self._scope_paint())
        self.extra_scope_rows.append(row)
        self._scope_only_open = False
        self._scope_build()
        try:
            row["w"]["name"].focus_set()
        except Exception:
            pass
        return row

    def _scope_line(self, section, item, row):
        desc, resp = row["desc"].get().strip().rstrip("."), row["resp"].get()
        return dict(section=section, item=item, desc=desc,
                    label=item + (f"  —  {desc}" if desc else ""),
                    resp=resp, bh=resp == "bh", customer=resp == "cust",
                    opt=resp == "opt", excl=False,
                    comment=row["comment"].get().strip(), print_comment=False)

    def _scope_out(self):
        """The lines as they are printed."""
        out = []
        for row in self.scope_rows:
            if row["included"].get() and not row["na"]:
                out.append(self._scope_line(row["item"]["section"],
                                            row["item"]["item"], row))
        for row in self.extra_scope_rows:
            name = row["label_var"].get().strip()
            if row["included"].get() and name:
                out.append(self._scope_line("OTHER", name, row))
        return out

    def _scope_state(self):
        """What the user decided, for the saved project file."""
        state = {}
        for row in self.scope_rows:
            s = dict(inc=row["included"].get(), resp=row["resp"].get(),
                     touched=row["touched"], comment=row["comment"].get())
            if row["desc"].get() != row["desc_auto"]:
                s["desc"] = row["desc"].get()
            state[row["item"]["key"]] = s
        extra = [dict(item=r["label_var"].get(), desc=r["desc"].get(),
                      resp=r["resp"].get(), inc=r["included"].get(),
                      comment=r["comment"].get())
                 for r in self.extra_scope_rows]
        return state, extra

    def _scope_restore(self, state, extra):
        for row in self.scope_rows:
            s = (state or {}).get(row["item"]["key"])
            if not s:
                continue
            row["included"].set(bool(s.get("inc", True)))
            row["touched"] = bool(s.get("touched"))
            if s.get("resp") in RESP_NAME and not row["forced"]:
                row["resp"].set(s["resp"])
            row["comment"].set(s.get("comment", ""))
            if "desc" in s:
                row["desc"].set(s["desc"])
        self.extra_scope_rows = []
        for e in extra or []:
            row = dict(label_var=tk.StringVar(value=e.get("item", "")),
                       included=tk.BooleanVar(value=bool(e.get("inc", True))),
                       resp=tk.StringVar(value=e.get("resp") if e.get("resp")
                                         in RESP_NAME else "bh"),
                       desc=tk.StringVar(value=e.get("desc", "")),
                       desc_auto="", comment=tk.StringVar(value=e.get("comment", "")),
                       touched=True, na=False, forced=None, w=None)
            row["label_var"].trace_add("write", lambda *_: self._scope_paint())
            self.extra_scope_rows.append(row)
        self._scope_build()

    # ── TAB 5: COMMERCIAL ─────────────────────────────────────────────────────
    def _tab_commercial(self):
        p = self.t_commercial
        self._comm_outer=p
        f=self._section(p,"Price & Commercial Conditions",
                        "applies to DTI and Budget offers")
        self._comm_body=f
        self._comm_card=f.shell
        g=tk.Frame(f,bg=BH_WHITE); g.pack(fill=tk.X)

        _lbl(g,"Base cost (€):",fnt=FONT_B).grid(row=0,column=0,sticky=tk.W,
                                                 padx=(0,14),pady=6)
        e=_entry(g,self.cost_base,w=22)
        e.grid(row=0,column=1,sticky=tk.W,padx=(0,30),pady=6)
        self._w_by_var[str(self.cost_base)]=e
        _lbl(g,"Contribution margin CM (%):",fnt=FONT_B).grid(
            row=0,column=2,sticky=tk.W,padx=(0,10))
        _entry(g,self.margin_pct,w=10).grid(row=0,column=3,sticky=tk.W,padx=(0,30))

        _lbl(g,"Warranty (%):",fnt=FONT_B).grid(row=1,column=0,sticky=tk.W,
                                                padx=(0,14),pady=6)
        _entry(g,self.warr_pct,w=10).grid(row=1,column=1,sticky=tk.W,padx=(0,30),pady=6)
        _lbl(g,"LDs (%):",fnt=FONT_B).grid(row=1,column=2,sticky=tk.W,padx=(0,10))
        _entry(g,self.lds_pct,w=10).grid(row=1,column=3,sticky=tk.W,padx=(0,30))

        _lbl(g,"Selling price (€):",fnt=FONT_B,fg=BH_GREEN).grid(
            row=2,column=0,sticky=tk.W,padx=(0,14),pady=(10,6))
        self._price_entry=tk.Entry(g,textvariable=self.final_price,width=22,
                 font=("Segoe UI",12,"bold"),bg=BH_WHITE,fg=BH_INK,relief=tk.FLAT,
                 bd=0,highlightthickness=1,highlightbackground=BH_BORDER,
                 highlightcolor=BH_GREEN,insertbackground=BH_GREEN)
        self._price_entry.grid(row=2,column=1,sticky=tk.W,padx=(0,30),pady=(10,6),ipady=3)
        Tooltip(self._price_entry,
                "Computed from the cost. Type a figure here to set the price "
                "yourself; the tool then shows the CM it leaves.")
        self._price_state=_lbl(g,"computed from the cost",fnt=FONT_SM,fg=BH_MUTED)
        self._price_state.grid(row=2,column=2,sticky=tk.W,padx=(0,10),pady=(10,6))
        _btn(g,"Recalculate from cost",lambda: self._set_price_mode("auto"),
             bg=BH_GRAY).grid(row=2,column=3,sticky=tk.W,pady=(10,6))

        note=tk.Frame(f,bg=BH_SEC,padx=12,pady=9,
                      highlightbackground=BH_BORDER,highlightthickness=1)
        note.pack(fill=tk.X,pady=(12,0))
        tk.Label(note,textvariable=self._price_note,font=FONT_B,bg=BH_SEC,
                 fg=BH_INK,anchor=tk.W,justify=tk.LEFT,
                 wraplength=1100).pack(anchor=tk.W)

        sep=tk.Frame(f,bg=BH_BORDER,height=1); sep.pack(fill=tk.X,pady=14)
        g2=tk.Frame(f,bg=BH_WHITE); g2.pack(fill=tk.X)
        _lbl(g2,"INCOTERM:",fnt=FONT_B).grid(row=0,column=0,sticky=tk.W,padx=(0,10))
        _combo(g2,self.incoterm,INCOTERMS,w=14).grid(row=0,column=1,sticky=tk.W,padx=(0,12))
        self._ic_tag=tk.Frame(g2,bg=BH_WHITE)
        self._ic_tag.grid(row=1,column=0,columnspan=4,sticky=tk.W,pady=(8,0))
        _lbl(g2,"Lead time:",fnt=FONT_B).grid(row=0,column=2,sticky=tk.W,padx=(0,10))
        e=_entry(g2,self.lead_time,w=30)
        e.grid(row=0,column=3,sticky=tk.W,padx=(0,30))
        self._w_by_var[str(self.lead_time)]=e

        for v in (self.cost_base,self.margin_pct,self.warr_pct,self.lds_pct):
            v.trace_add("write",self._calc_price)
        self.final_price.trace_add("write",self._price_typed)
        self._set_price_mode("auto")

        # ── the selected rule, drawn ──
        f2=self._section(p,"Delivery Terms Explained",
                         "who pays what, and where the risk passes")
        self._comm_card2=f2.shell
        self._ic_canvas=tk.Canvas(f2,bg=BH_WHITE,height=172,
                                  highlightthickness=0,bd=0)
        self._ic_canvas.pack(fill=tk.X)
        self._ic_note=tk.StringVar()
        nb=tk.Frame(f2,bg=BH_SEC,padx=12,pady=9,
                    highlightbackground=BH_BORDER,highlightthickness=1)
        nb.pack(fill=tk.X,pady=(8,0))
        tk.Label(nb,textvariable=self._ic_note,font=FONT_N,bg=BH_SEC,fg=BH_INK,
                 justify=tk.LEFT,wraplength=1100,anchor=tk.W).pack(anchor=tk.W)

        f3=self._section(p,"All Incoterms 2020 at a Glance",
                         "the selected rule is highlighted")
        self._comm_card3=f3.shell
        self._ic_matrix=tk.Canvas(f3,bg=BH_WHITE,height=232,
                                  highlightthickness=0,bd=0)
        self._ic_matrix.pack(fill=tk.X)

        self._ic_canvas.bind("<Configure>", lambda e: self._draw_incoterm())
        self._ic_matrix.bind("<Configure>", lambda e: self._draw_incoterm())
        self.incoterm.trace_add("write", lambda *a: self._draw_incoterm())
        self._draw_incoterm()

    def _draw_incoterm(self, *_):
        if not hasattr(self, "_ic_canvas"):
            return
        code = self.incoterm.get()
        c1, c2 = self._ic_canvas, self._ic_matrix
        draw_incoterm_tk(c1, code, max(c1.winfo_width(), 900), 172)
        draw_incoterm_matrix_tk(c2, code, max(c2.winfo_width(), 900), 232)
        sp = incoterm_spec(code)
        if sp:
            note = incoterm_note(code)
            self._ic_note.set(f"{code} — {sp['name']}.  {sp['note']}"
                              + (f"\n\n{note}" if note else ""))
        tag = getattr(self, "_ic_tag", None)
        if tag is not None:
            for w in tag.winfo_children():
                w.destroy()
            if code in INCOTERMS_RECOMMENDED:
                _pill(tag, "Recommended", BH_OK).pack(side=tk.LEFT)
            elif code in INSURANCE_NOTE_RULES:
                _pill(tag, "Insurance to be settled in the sales contract",
                      BH_WARN).pack(side=tk.LEFT)
                _lbl(tag, "   the note is printed under the Incoterm",
                     fnt=FONT_SM, fg=BH_MUTED).pack(side=tk.LEFT)

    # ── ACTION BAR ────────────────────────────────────────────────────────────
    def _action_bar(self):
        wrap = tk.Frame(self.root, bg=BH_DARK)
        wrap.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Frame(wrap, bg=BH_LINE_D, height=1).pack(fill=tk.X)
        bar = tk.Frame(wrap, bg=BH_DARK, height=64)
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=BH_DARK)
        left.pack(side=tk.LEFT, padx=(20, 0), pady=13)
        cta = _SoftButton(left, "Generate PDF", self._generate_pdf,
                          fill=BH_GREEN, host_bg=BH_DARK, fg=BH_WHITE,
                          fnt=FONT_B, pad=(22, 9), radius=R_SM)
        cta.pack(side=tk.LEFT)
        self._gen_buttons = getattr(self, "_gen_buttons", []) + [cta]
        Tooltip(cta, "Assemble the complete technical proposal   (Ctrl+G)")
        for txt, cmd, key in [("Save",            self._save_json,    "Ctrl+S"),
                              ("Open",            self._load_json,    "Ctrl+O"),
                              ("New offer",       self._reset,        "Ctrl+N"),
                              ("Export scope to Excel", self._export_excel, "Ctrl+E")]:
            b = _SoftButton(left, txt, cmd, fill=BH_DARK, host_bg=BH_DARK,
                            fg=BH_ON_DARK2, fnt=FONT_B, pad=(16, 9),
                            radius=R_SM, border=BH_LINE_D)
            b.pack(side=tk.LEFT, padx=(9, 0))
            Tooltip(b, f"{txt}   ({key})" if key else txt)

        self._status = _StatusVar(self.root, value="Ready")
        sf = tk.Frame(bar, bg=BH_DARK)
        sf.pack(side=tk.RIGHT, padx=24)
        tk.Frame(sf, bg=BH_GREEN, width=7, height=7).pack(side=tk.LEFT,
                                                          padx=(0, 9))
        tk.Label(sf, textvariable=self._status, font=FONT_SM,
                 bg=BH_DARK, fg=BH_ON_DARK2).pack(side=tk.LEFT)

    # ─────────────────────────────────────────────────────────────────────────
    # CALLBACKS
    # ─────────────────────────────────────────────────────────────────────────
    def _toggle_gas(self):
        self._gas_entry.config(
            state=tk.NORMAL if self.gas_type.get()=="Custom" else tk.DISABLED)

    def _update_nda(self):
        if self.nda_signed.get():
            self._nda_lbl.config(text="✓  NDA Signed",fg=BH_GREEN)
            note = "The datasheet is inserted in full."
        else:
            self._nda_lbl.config(text="✗  Not Signed",fg=BH_RED)
            note = ("The datasheet stops after the gas composition table: the "
                    "sheets that follow it are left out of the proposal.")
        if hasattr(self, "_nda_note"):
            self._nda_note.config(text=note)

    def _on_phase(self, _=None):
        self.back_to_back.set("Back to Back" if self.num_phases.get()=="2" else "Simple")

    def _on_category(self):
        cat=self.offer_category.get()
        if cat != getattr(self, "_lib_cat", cat) or not getattr(self, "_lib_started", False):
            self._lib_started = True
            self._lib_load_category(cat)
            self._lib_paint()
        if self._project_pick:
            root = proposal_root(self._project_pick, cat)
            if root != self.user_root.get():
                self.user_root.set(root)
        self._paint_category()
        for row in self.scope_rows:
            if not row["touched"] and not row["forced"]:
                row["resp"].set(row["item"]["resp"].get(cat, "bh"))
        self._scope_paint()
        if cat=="FIRM":
            for nm in ("_comm_card","_comm_card2","_comm_card3"):
                if hasattr(self,nm): getattr(self,nm).pack_forget()
            if not hasattr(self,"_firm_notice"):
                self._firm_notice=tk.Frame(self._comm_outer,bg=BH_BG)
                tk.Label(self._firm_notice,
                         text="Pricing for firm offers is handled in the formal "
                              "quotation process, not in this tool.",
                         font=FONT_N,bg=BH_BG,fg=BH_GRAY,justify=tk.CENTER,pady=34).pack()
            self._firm_notice.pack(fill=tk.BOTH,expand=True)
        else:
            if hasattr(self,"_firm_notice"): self._firm_notice.pack_forget()
            for nm in ("_comm_card","_comm_card2","_comm_card3"):
                if hasattr(self,nm):
                    getattr(self,nm).pack(fill=tk.X,padx=16,pady=(0,12))
            self._draw_incoterm()
        if hasattr(self,"_doc_container"): self._refresh_docs()

    def _on_plc(self, _=None):
        docs=PLC_DOCS.get(self.ucp_type.get(),{}).get(self.plc_option.get(),[])
        self._plc_doc_lbl.config(text="  ·  ".join(docs))
        if hasattr(self,"_plc_doc_container"): self._refresh_plc_container()
        if hasattr(self,"_scope_frame"):       self._render_scope()

    def _concept_changed(self, *_):
        if self._sync_guard:
            return
        c = self.cooling_system.get()
        if c not in COOLING_CONCEPTS:
            return
        # the motor cooling written in the scope of supply is read from the
        # concept; the concept itself is never changed by the tool
        f = _concept_features(c)
        want = "Fan" if f["fan"] else "TCV" if f["tcv"] else ""
        if want and self.motor_cooling.get() != want:
            self.motor_cooling.set(want)

    def _auto_code(self, *_):
        ni = self.num_impellers.get().strip()
        cs = _casing_digits(self.casing_size.get())
        ph = self.num_phases.get().strip()
        code = ""
        if ni and cs and ph:
            code = f"2ICL{cs}{ni}" if ph == "2" else f"ICL{cs}{ni}"
        self._code_auto = code
        if not self.code_manual.get() and self.config_code.get() != code:
            self._code_guard = True
            self.config_code.set(code)
            self._code_guard = False
        self._code_hint_paint()
        want = "Back to Back" if ph == "2" else "Simple"
        if ph in ("1", "2") and self.back_to_back.get() != want:
            self.back_to_back.set(want)

    def _code_typed(self, *_):
        """Anything typed in the name that differs from the automatic one
        makes it a forced name; typing the automatic one back releases it."""
        if getattr(self, "_code_guard", False):
            return
        auto = getattr(self, "_code_auto", "")
        typed = self.config_code.get().strip()
        manual = bool(typed) and typed != auto
        if self.code_manual.get() != manual:
            self.code_manual.set(manual)
        self._code_hint_paint()

    def _code_to_auto(self):
        self.code_manual.set(False)
        self._auto_code()
        self._status.set("Compressor name follows the configuration again.")

    def _code_hint_paint(self):
        lb = getattr(self, "_code_hint", None)
        if lb is None:
            return
        try:
            auto = getattr(self, "_code_auto", "")
            if self.code_manual.get():
                lb.config(text=f"forced by hand  ·  configuration gives {auto or '—'}",
                          fg=BH_WARN)
                self._code_reset.pack(side=tk.LEFT, padx=(10, 0))
                self._code_entry.config(fg=BH_WARN)
            else:
                lb.config(text="follows impellers, casing and phases  ·  "
                               "type over it to force a name", fg=BH_MUTED)
                self._code_reset.pack_forget()
                self._code_entry.config(fg=BH_GREEN)
        except tk.TclError:
            pass

    def _gen_code(self):
        ni=self.num_impellers.get().strip()
        cs=_casing_digits(self.casing_size.get())
        np_=self.num_phases.get().strip()
        if not all([ni,cs,np_]):
            messagebox.showwarning("Incomplete","Fill all configuration fields."); return
        code=f"{np_}ICL{cs}{ni}" if np_=="2" else f"ICL{cs}{ni}"
        self.config_code.set(code)
        self._status.set(f"Config code: {code}")
        self._render_scope()

    def _calc_price(self, *_):
        """Cost to price, the way the ICL budget sheet does it:

            price = cost / (1 - warranty - LDs - CM), rounded up to 10 €

        Warranty, LDs and CM are shares of the selling price, not of the
        cost. When the price is typed by hand it is kept as it is and the
        CM it leaves is worked back from it."""
        cost = _num(self.cost_base.get())
        w = _num(self.warr_pct.get()) / 100.0
        l = _num(self.lds_pct.get()) / 100.0
        cm = _num(self.margin_pct.get()) / 100.0

        if self.price_mode.get() == "manual":
            price = _num(self.final_price.get())
            if price <= 0:
                self._price_note.set("Type the selling price, or recalculate "
                                     "it from the cost.")
                return
            if cost > 0:
                left = (1 - w - l - cost / price) * 100
                self._price_note.set(
                    f"Price entered by hand {price:,.0f} €      cost {cost:,.0f} €      "
                    f"warranty {w*price:,.0f} €      LDs {l*price:,.0f} €      "
                    f"CM left {left:.1f} %")
            else:
                self._price_note.set(f"Price entered by hand {price:,.0f} €")
            return

        if cost <= 0:
            self._price_note.set("Enter a base cost to compute the price.")
            return
        denom = 1 - w - l - cm
        if denom <= 0.001:
            self._price_note.set("Warranty + LDs + CM reach 100 % of the price: "
                                 "no selling price can cover that.")
            return
        price = math.ceil(round(cost / denom, 6) / 10) * 10
        self._price_busy = True
        try:
            self.final_price.set(f"{price:,.0f}")
        finally:
            self._price_busy = False
        self._price_note.set(
            f"{cost:,.0f} € / (1 − {w*100:.3g} % − {l*100:.3g} % − {cm*100:.3g} %)"
            f"  =  {price:,.0f} €      CM {cm*price:,.0f} €      "
            f"warranty {w*price:,.0f} €      LDs {l*price:,.0f} €")
        if hasattr(self, "_status"):
            self._status.set(f"Selling price: {price:,.0f} €", toast=False)

    def _price_typed(self, *_):
        """Any edit of the price field made by the user takes over."""
        if self._price_busy:
            return
        if self.price_mode.get() != "manual":
            self.price_mode.set("manual")
            if hasattr(self, "_price_state"):
                self._price_state.config(text="entered by hand", fg=BH_WARN)
        self._calc_price()

    def _set_price_mode(self, mode):
        mode = "manual" if mode == "manual" else "auto"
        self.price_mode.set(mode)
        if hasattr(self, "_price_state"):
            self._price_state.config(
                text="entered by hand" if mode == "manual"
                else "computed from the cost",
                fg=BH_WARN if mode == "manual" else BH_MUTED)
        self._calc_price()

    # ─────────────────────────────────────────────────────────────────────────
    # DATA COLLECTION
    # ─────────────────────────────────────────────────────────────────────────
    def _collect(self):
        plc=self.ucp_type.get(); opt=self.plc_option.get()
        scope_out=self._scope_out()
        scope_state, scope_extra = self._scope_state()
        return dict(
            offer_category=self.offer_category.get(),
            opp_id=self.opp_id.get(), revision=self.revision.get(),
            comment=self.gen_comment.get(),
            application=self.application.get(),
            gas_type=self.gas_type.get(), custom_gas=self.custom_gas.get(),
            client_name=self.client_name.get(), client_loc=self.client_loc.get(),
            project_date=self.project_date.get(),
            ae_name=self.ae_name.get(), ae_phone=self.ae_phone.get(),
            ae_email=self.ae_email.get(),
            sales_name=self.sales_name.get(), sales_phone=self.sales_phone.get(),
            sales_email=self.sales_email.get(),
            adv_name=self.adv_name.get(), adv_phone=self.adv_phone.get(),
            adv_email=self.adv_email.get(),
            sales_title=self.sales_title.get(), adv_title=self.adv_title.get(),
            ae_title=self.ae_title.get(),
            lib_state=self._lib_state_now(), lib_choice=dict(self.lib_choice),
            project_pick=self._project_pick,
            motor_cooling=self.motor_cooling.get(),
            scope_state=scope_state, scope_extra=scope_extra,
            nda_signed=self.nda_signed.get(),
            documents=[d["name"] for d in DB_LIBRARY if self._lib_on(d["id"])],
            num_impellers=self.num_impellers.get(), casing_size=self.casing_size.get(),
            num_phases=self.num_phases.get(), config_code=self.config_code.get(),
            code_manual=self.code_manual.get(),
            motor_type=self.motor_type.get(), back_to_back=self.back_to_back.get(),
            rating=self.rating.get(), ucp_type=plc, plc_option=opt,
            plc_documents=PLC_DOCS.get(plc,{}).get(opt,[]),
            plc_safety=self.plc_safety.get(), thrust_bearing=self.thrust_bearing.get(),
            motor_cooling_tcv=self.motor_cooling_tcv.get(),
            motor_cooling_fan=self.motor_cooling_fan.get(),
            vfd_cooling=self.vfd_cooling.get(), transfo_type=self.transfo_type.get(),
            power=self.power.get(), cooling_system=self.cooling_system.get(),
            antisurge_valve=self.antisurge_valve.get(), filter_type=self.filter_type.get(),
            ehouse=self.ehouse.get(), mcc=self.mcc.get(),
            comp_enclosure=self.comp_enclosure.get(),
            clarifications=self._clarif.get("1.0",tk.END).strip(),
            cost_base=self.cost_base.get(), margin_pct=self.margin_pct.get(),
            warr_pct=self.warr_pct.get(), lds_pct=self.lds_pct.get(),
            price_mode=self.price_mode.get(),
            project_name=self.project_name.get(), nb_units=self.nb_units.get(),
            pdf_frame=self.pdf_frame.get(), skeleton_file=self.skeleton_file.get(),
            doc_attach=self.doc_attach, doc_layout=self.doc_layout,
            pdf_output=self.pdf_output.get(), pdf_use_word=self.pdf_use_word.get(),
            final_price=self.final_price.get(), incoterm=self.incoterm.get(),
            lead_time=self.lead_time.get(),
            scope_of_supply=scope_out,
            # ── document management ──────────────────────────────────────
            user_root=self.user_root.get(), db_root=self.db_root.get(),
            use_db=self.use_db.get(), auto_place=self.auto_place.get(),
            doc_use_db={k: v.get() for k, v in self.doc_use_db.items()},
            document_structure=[
                dict(code=it["ref"], section=sec["num"], depth=depth,
                     section_title=sec["title"], folder=sec["folder"],
                     title=self._doc_title(it), origin=it["src"],
                     status=self.doc_status.get(it["code"], {}).get("label") or "—",
                     source=self.doc_status.get(it["code"], {}).get("src", ""),
                     file=self.doc_status.get(it["code"], {}).get("path", ""))
                for sec in self._doc_struct() for it, depth in _leaves(sec)],
        )

    # ─────────────────────────────────────────────────────────────────────────
    # EXCEL EXPORT — SCOPE OF SUPPLY
    # ─────────────────────────────────────────────────────────────────────────
    def _export_excel(self):
        data=self._collect()
        path=filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook","*.xlsx")],
            initialfile=(f"ScopeOfSupply_{data['opp_id'] or 'NEW'}_"
                         f"{data['revision'] or 'R0'}.xlsx"),
            title="Export Scope of Supply to Excel")
        if not path: return
        try:
            self._write_excel(data,path)
            self._status.set(f"Exported: {os.path.basename(path)}")
            messagebox.showinfo("Export Complete",f"Scope of Supply exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error",str(e))

    def _write_excel(self, data, path):
        wb=openpyxl.Workbook()
        ws=wb.active
        ws.title="Scope of Supply"
        ws.sheet_view.showGridLines=False

        cat_hex={"DTI":"00D4C5","BUDGET":"FFC107","FIRM":"DC3545"}.get(
            data["offer_category"],"018374")

        def Fl(c): return PatternFill("solid",fgColor=c)
        def Ft(sz=9,bold=False,color="1E2423",italic=False):
            return Font(name="Segoe UI",size=sz,bold=bold,color=color,italic=italic)
        def FtC(sz=9,bold=False,color="1E2423"):
            return Font(name="Courier New",size=sz,bold=bold,color=color)

        thin=Side(style="thin",color="DCE0DE")
        bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
        al_l=Alignment(horizontal="left", vertical="center",wrap_text=True)
        al_c=Alignment(horizontal="center",vertical="center",wrap_text=True)
        al_r=Alignment(horizontal="right", vertical="center")

        NC=6
        for col,width in zip("ABCDEF",(5,46,62,13,13,11)):
            ws.column_dimensions[col].width=width

        r=1
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
        c=ws.cell(r,1,"BAKER HUGHES THERMODYN  —  DTI DEPARTMENT")
        c.font=Ft(14,True,"FFFFFF"); c.fill=Fl("05322B"); c.alignment=al_c
        ws.row_dimensions[r].height=36; r+=1

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
        c=ws.cell(r,1,f"SCOPE OF SUPPLY  |  Offer Category: {data['offer_category']}")
        c.font=Ft(11,True,"FFFFFF"); c.fill=Fl("018374"); c.alignment=al_c
        ws.row_dimensions[r].height=24; r+=1

        meta=(f"OPP ID: {data['opp_id'] or '—'}    "
              f"Rev.: {data['revision'] or '—'}    "
              f"Client: {data['client_name'] or '—'}    "
              f"Date: {data['project_date']}    "
              f"INCOTERM: {data['incoterm']}")
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
        c=ws.cell(r,1,meta)
        c.font=Ft(8,italic=True,color="4E5755"); c.fill=Fl("F4F5F4"); c.alignment=al_l
        ws.row_dimensions[r].height=16; r+=2

        for ci,(txt,al) in enumerate(zip(
            ["#","Item","Description","Baker Hughes","Customer","Option"],
            [al_c,al_l,al_l,al_c,al_c,al_c]),1):
            ce=ws.cell(r,ci,txt)
            ce.font=Ft(9,True,"FFFFFF"); ce.fill=Fl("018374")
            ce.alignment=al; ce.border=bdr
        ws.row_dimensions[r].height=20; r+=1

        cur_sec=None; idx=1
        for item in data["scope_of_supply"]:
            sec=item["section"]
            if sec and sec!=cur_sec:
                cur_sec=sec
                ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
                ce=ws.cell(r,1,f"  {sec}")
                ce.font=Ft(9,True,"FFFFFF"); ce.fill=Fl("05322B"); ce.alignment=al_l
                ws.row_dimensions[r].height=18; r+=1

            bg=Fl("FFFFFF") if idx%2 else Fl("F6F8F7")
            resp=item.get("resp","")
            cells=[(idx,Ft(8,color="7B8482"),al_c),
                   (item.get("item") or item["label"],Ft(9),al_l),
                   (item.get("desc",""),Ft(8,color="4E5755"),al_l),
                   ("X" if resp=="bh" else "",Ft(11,True,"018374"),al_c),
                   ("X" if resp=="cust" else "",Ft(11,True,"1E2423"),al_c),
                   ("X" if resp=="opt" else "",Ft(11,True,"B8860B"),al_c)]
            for ci,(val,font,al) in enumerate(cells,1):
                ce=ws.cell(r,ci,val)
                ce.font=font; ce.fill=bg; ce.alignment=al; ce.border=bdr
            long=max(len(str(item.get("desc",""))),len(str(item.get("item",""))))
            ws.row_dimensions[r].height=22 if long<70 else 34
            r+=1; idx+=1

        # Footer
        r+=1
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
        ce=ws.cell(r,1,
                   f"Generated by {APP_NAME} · {APP_SUB}  —  "
                   f"{datetime.now().strftime('%Y-%m-%d  %H:%M')}")
        ce.font=Ft(7,italic=True,color="7B8482"); ce.fill=Fl("F4F5F4"); ce.alignment=al_r
        ws.row_dimensions[r].height=14

        ws.freeze_panes=ws.cell(row=6,column=1)
        ws.page_setup.orientation="landscape"
        ws.page_setup.fitToPage=True
        ws.page_setup.fitToWidth=1
        ws.print_title_rows="1:5"

        notes=[it for it in data["scope_of_supply"] if (it.get("comment") or "").strip()]
        if notes:
            wn=wb.create_sheet("Internal notes")
            wn.sheet_view.showGridLines=False
            wn.sheet_properties.tabColor="B8860B"
            for col,wd in zip("ABCD",(22,46,14,60)):
                wn.column_dimensions[col].width=wd
            wn.merge_cells("A1:D1")
            c=wn.cell(1,1,"INTERNAL — Baker Hughes only. Delete this sheet before "
                          "sending the file to the customer.")
            c.font=Ft(10,True,"FFFFFF"); c.fill=Fl("B8860B"); c.alignment=al_l
            wn.row_dimensions[1].height=24
            for ci,t in enumerate(("Section","Item","Supplier","Internal note"),1):
                ce=wn.cell(3,ci,t); ce.font=Ft(9,True,"FFFFFF"); ce.fill=Fl("018374")
                ce.border=bdr; ce.alignment=al_l
            for ri,it in enumerate(notes,4):
                for ci,v in enumerate((it.get("section",""),it.get("item") or it.get("label",""),
                                       {"bh":"Baker Hughes","cust":"Customer","opt":"Option"}
                                       .get(it.get("resp"),""),it["comment"]),1):
                    ce=wn.cell(ri,ci,v); ce.font=Ft(9); ce.border=bdr; ce.alignment=al_l
        self._write_doc_sheet(wb,data,Fl,Ft,bdr,al_l,al_c)
        wb.save(path)

    # ── second sheet: document package status ─────────────────────────────────
    def _write_doc_sheet(self, wb, data, Fl, Ft, bdr, al_l, al_c):
        ws=wb.create_sheet("Documents")
        ws.sheet_view.showGridLines=False
        for col,wd in zip("ABCDEF",(9,46,20,16,34,44)):
            ws.column_dimensions[col].width=wd

        r=1
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
        c=ws.cell(r,1,f"TECHNICAL PROPOSAL DOCUMENTS  |  {data['offer_category']} OFFER")
        c.font=Ft(12,True,"FFFFFF"); c.fill=Fl("05322B"); c.alignment=al_c
        ws.row_dimensions[r].height=28; r+=1

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
        c=ws.cell(r,1,f"Project folder: {data['user_root'] or '—'}    "
                      f"Database: {data['db_root'] or '—'}    "
                      f"Database in use: {'Yes' if data['use_db'] else 'No'}")
        c.font=Ft(8,italic=True,color="4E5755"); c.fill=Fl("F4F5F4"); c.alignment=al_l
        ws.row_dimensions[r].height=16; r+=2

        for ci,txt in enumerate(["#","Document","Status","Source","Section","File"],1):
            ce=ws.cell(r,ci,txt)
            ce.font=Ft(9,True,"FFFFFF"); ce.fill=Fl("018374")
            ce.alignment=al_c if ci in (1,3,4) else al_l; ce.border=bdr
        ws.row_dimensions[r].height=20; r+=1

        cur=None
        colours={"In the DB":("E7F1EA","1A6B3C"),
                 "Provided":("E7F1EA","018374"),
                 "Need to be provided":("F7EFDC","B8860B"),
                 "Not found in the DB":("F6E7E7","B22222"),
                 "Built by the tool":("EDF0EF","3D4745")}
        for d in data.get("document_structure",[]):
            if d["section"]!=cur:
                cur=d["section"]
                ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
                ce=ws.cell(r,1,f"  SECTION {d['section']} — {d['section_title']}")
                ce.font=Ft(9,True,"FFFFFF"); ce.fill=Fl("05322B"); ce.alignment=al_l
                ws.row_dimensions[r].height=18; r+=1
            bgc,fgc=colours.get(d["status"],("FFFFFF","1E2423"))
            vals=[d["code"],("      " * d.get("depth",0))+d["title"],
                  d["status"],d["source"],d["folder"],
                  os.path.basename(d["file"]) if d["file"] else ""]
            for ci,v in enumerate(vals,1):
                ce=ws.cell(r,ci,v)
                ce.fill=Fl(bgc); ce.border=bdr
                ce.alignment=al_c if ci in (1,3,4) else al_l
                ce.font=(Ft(9,True,fgc) if ci==3 else
                         Ft(8,color="4E5755") if ci in (5,6) else Ft(9))
            ws.row_dimensions[r].height=18; r+=1

        ws.freeze_panes=ws.cell(row=5,column=1)
        ws.page_setup.orientation="landscape"
        ws.page_setup.fitToPage=True
        ws.page_setup.fitToWidth=1

    # ─────────────────────────────────────────────────────────────────────────
    # SAVE / LOAD
    # ─────────────────────────────────────────────────────────────────────────
    # ── readiness ─────────────────────────────────────────────────────────────
    def _readiness(self):
        """Everything a customer would see as unfinished, page by page."""
        out = []

        def need(page, text, var=None, focus=None):
            out.append(dict(page=page, text=text, var=var, focus=focus))
        g = lambda v: v.get().strip()
        if not g(self.client_name):
            need(0, "Customer name", self.client_name)
        if not g(self.project_name):
            need(0, "Project name", self.project_name)
        if not g(self.opp_id):
            need(0, "Proposal reference (OPP ID)", self.opp_id)
        if not g(self.sales_name):
            need(0, "Sales manager", self.sales_name)
        if not g(self.ae_name):
            need(0, "Application engineer", self.ae_name)
        if not g(self.config_code):
            need(1, "Compressor configuration", self.num_impellers)
        if self.cooling_system.get() not in COOLING_CONCEPTS:
            need(1, "Cooling concept", self.cooling_system)
        miss = [d["name"] for d in DB_LIBRARY
                if self._lib_on(d["id"]) and d["rule"] != "tool"
                and (self._lib_res.get(d["id"]) or {}).get("quality") == "missing"]
        if miss:
            need(2, f"{len(miss)} ticked document(s) not found: "
                    + ", ".join(miss[:3]) + (" …" if len(miss) > 3 else ""))
        n = self._scope_open_count()
        if n:
            need(3, f"{n} scope line(s) to confirm", focus="scope")
        sus = self._scope_suspects()
        if sus:
            need(3, f"{len(sus)} scope description(s) read like an internal note "
                    f"(“{sus[0][1]}”): the customer will read them", focus="scope")
        if self.offer_category.get() != "FIRM":
            if _num(self.final_price.get()) <= 0:
                need(4, "Selling price", self.cost_base)
            if not g(self.lead_time):
                need(4, "Lead time", self.lead_time)
        if not g(self.user_root):
            need(5, "Project folder", self.user_root)
        else:
            todo = [st.get("title", "") for st in self.doc_status.values()
                    if st.get("state") == "need"]
            if todo:
                need(5, f"{len(todo)} project document(s) to provide: "
                        + ", ".join(todo[:3]) + (" …" if len(todo) > 3 else ""))
        return out

    def _ux_tick(self):
        """Every second and a half: badges, highlighted fields, the panel of
        what is left, the title bar, the automatic save."""
        try:
            items = self._readiness()
            self._ready_paint(items)
            self._autosave()
            self._title_paint()
        except tk.TclError:
            return
        except Exception:
            pass
        self._tick_id = self.root.after(1500, self._ux_tick)

    def _ready_paint(self, items):
        per = {}
        for it in items:
            per[it["page"]] = per.get(it["page"], 0) + 1
        for pg in getattr(self, "_pages_list", []):
            n = per.get(pg["idx"], 0)
            pg["badge"].config(text=str(n) if n else "✓",
                               fg=BH_WARN if n else BH_ACCENT)
        flagged = {str(it["var"]) for it in items if it.get("var") is not None}
        focus = self.root.focus_get()
        for name, w in self._w_by_var.items():
            try:
                if w is focus:
                    continue
                if name in flagged:
                    w.config(highlightbackground=BH_WARN) if isinstance(w, tk.Entry) \
                        else None
                elif isinstance(w, tk.Entry) and \
                        str(w.cget("highlightbackground")).upper() == BH_WARN.upper():
                    w.config(highlightbackground=BH_BORDER)
            except tk.TclError:
                pass
        sig = tuple((it["page"], it["text"]) for it in items)
        if sig == self._ready_sig or not hasattr(self, "_ready_box"):
            return
        self._ready_sig = sig
        box = self._ready_box
        for w in box.winfo_children():
            w.destroy()
        if not items:
            _lbl(box, "✓  Everything is filled in: the proposal is ready to go out.",
                 fnt=FONT_B, fg=BH_ACCENT, bg=BH_INK).pack(anchor=tk.W)
            return
        _lbl(box, f"{len(items)} point(s) to settle before sending it — click to go there",
             fnt=FONT_B, fg=BH_WHITE, bg=BH_INK).pack(anchor=tk.W, pady=(0, 4))
        for it in items[:7]:
            page = self._pages_list[it["page"]]["name"]
            ln = tk.Label(box, text=f"•  {it['text']}      {page}",
                          font=FONT_SM, bg=BH_INK, fg=BH_ON_DARK2,
                          cursor="hand2", anchor=tk.W)
            ln.pack(anchor=tk.W)
            ln.bind("<Button-1>", lambda e, i=it: self._jump(i))
            ln.bind("<Enter>", lambda e, w=ln: w.config(fg=BH_ACCENT))
            ln.bind("<Leave>", lambda e, w=ln: w.config(fg=BH_ON_DARK2))
        if len(items) > 7:
            _lbl(box, f"   and {len(items) - 7} more", fnt=FONT_SM,
                 fg=BH_ON_DARK, bg=BH_INK).pack(anchor=tk.W)

    def _jump(self, item):
        self._show_page(item["page"])
        if item.get("focus") == "scope":
            self._scope_only_open = True
            self._scope_build()
        w = self._w_by_var.get(str(item["var"])) if item.get("var") is not None else None
        if w is not None:
            self.root.after(260, lambda: self._reveal(w))

    def _reveal(self, w):
        try:
            inner = self._pages_list[self._page_idx]["inner"]
            cv = inner._canvas
            y = w.winfo_rooty() - inner.winfo_rooty()
            cv.yview_moveto(max(0.0, (y - 140) / max(inner.winfo_height(), 1)))
            w.focus_set()
            if isinstance(w, tk.Entry):
                w.config(highlightbackground=BH_WARN, highlightcolor=BH_WARN)
                w.after(1600, lambda: w.config(highlightcolor=BH_GREEN))
        except tk.TclError:
            pass

    # ── previous / next ──────────────────────────────────────────────────────
    def _page_nav(self):
        pages = self._pages_list
        for i, pg in enumerate(pages):
            nav = tk.Frame(pg["inner"], bg=BH_BG)
            nav.pack(fill=tk.X, padx=SP4, pady=(SP3, SP5))
            tk.Frame(nav, bg=BH_BORDER, height=1).pack(fill=tk.X, pady=(0, 14))
            row = tk.Frame(nav, bg=BH_BG); row.pack(fill=tk.X)
            if i > 0:
                _btn(row, f"←  {pages[i - 1]['name']}",
                     lambda k=i - 1: self._show_page(k), bg=BH_GRAY).pack(side=tk.LEFT)
            if i < len(pages) - 1:
                _btn(row, f"Next:  {pages[i + 1]['name']}  →",
                     lambda k=i + 1: self._show_page(k)).pack(side=tk.RIGHT)
            else:
                _btn(row, "Generate the proposal", self._generate_pdf).pack(side=tk.RIGHT)

    # ── automatic save and resume ────────────────────────────────────────────
    def _autosave_path(self):
        base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
        folder = os.path.join(base, "Y-B-TOOL") if base else _tool_dir()
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            folder = _tool_dir()
        return os.path.join(folder, "y_b_tool_autosave.json")

    def _autosave(self):
        if not self._autosave_armed:
            return
        data = self._collect()
        key = self._content_key(data)
        if key == self._autosave_key:
            return
        self._autosave_key = key
        if not any((data.get(k) or "").strip()
                   for k in ("client_name", "opp_id", "project_name")):
            return
        data["_file"] = self._file
        data["_saved_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        try:
            tmp = self._autosave_path() + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False)
            os.replace(tmp, self._autosave_path())
        except OSError:
            pass

    def _resume_offer(self):
        """The offer being edited when the tool was last closed is offered
        back, without a dialog in the way."""
        try:
            with open(self._autosave_path(), encoding="utf-8") as fh:
                d = json.load(fh)
        except Exception:
            self._autosave_armed = True
            return
        who = "  ·  ".join(x for x in (d.get("client_name"), d.get("project_name"),
                                       f"OPP {d['opp_id']}" if d.get("opp_id") else "")
                          if x)
        if not who:
            self._autosave_armed = True
            return
        inner = self._pages_list[0]["inner"]
        bar = tk.Frame(inner, bg=BH_NAVY, padx=18, pady=12)
        first = inner.winfo_children()[0]
        bar.pack(fill=tk.X, padx=SP4, pady=(SP3, 0), before=first)
        _lbl(bar, "Offer in progress", fnt=FONT_B, fg=BH_WHITE, bg=BH_NAVY).pack(
            side=tk.LEFT)
        _lbl(bar, f"   {who}      last change {d.get('_saved_at', '')}",
             fnt=FONT_SM, fg=BH_ON_DARK2, bg=BH_NAVY).pack(side=tk.LEFT)
        _btn(bar, "Start a new offer", self._dismiss_resume, bg=BH_NAVY_HI).pack(
            side=tk.RIGHT)
        _btn(bar, "Resume", lambda: self._resume(d)).pack(side=tk.RIGHT, padx=8)
        self._resume_bar = bar

    def _resume(self, d):
        file = d.pop("_file", "") or ""
        d.pop("_saved_at", None)
        self._apply(d)
        self._file = file if file and os.path.exists(file) else ""
        self._saved_key = None
        self._dismiss_resume()
        self._status.set("Offer resumed where you left it.")

    def _dismiss_resume(self):
        bar = getattr(self, "_resume_bar", None)
        if bar is not None:
            try:
                bar.destroy()
            except tk.TclError:
                pass
            self._resume_bar = None
        self._autosave_armed = True

    def _on_close(self):
        if getattr(self, "_gen_busy", False) and not messagebox.askyesno(
                "Proposal in progress",
                "A proposal is still being generated.\n\nClose the tool and "
                "stop it?"):
            return
        if self._dirty() and self._file:
            if not self._confirm_discard():
                return
        else:
            self._autosave()            # kept for next time, nothing lost
        if self._tick_id:
            try:
                self.root.after_cancel(self._tick_id)
            except Exception:
                pass
        self.root.destroy()

    def _content_key(self, data=None):
        data = self._collect() if data is None else data
        return json.dumps(data, sort_keys=True, default=str)

    def _save_json(self, ask=False):
        self._remember_contacts()
        data = self._collect()
        path = self._file
        if ask or not path:
            path = filedialog.asksaveasfilename(
                title="Save the offer", defaultextension=".json",
                filetypes=[("Offer", "*.json")],
                initialdir=self.user_root.get().strip() or _tool_dir(),
                initialfile=f"offer_{data['opp_id'] or 'NEW'}"
                            f"{'_rev' + data['revision'] if data.get('revision') else ''}.json")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except OSError as e:
            messagebox.showerror("Save", str(e)); return
        self._file = path
        self._saved_key = self._content_key(data)
        self._title_paint()
        self._status.set(f"Saved — {os.path.basename(path)}")

    def _load_json(self, path=None):
        if not self._confirm_discard():
            return
        path = path or filedialog.askopenfilename(
            title="Open an offer", filetypes=[("Offer", "*.json")],
            initialdir=self.user_root.get().strip() or _tool_dir())
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                d = json.load(f)
            self._apply(d)
        except Exception as e:
            messagebox.showerror("Open", str(e)); return
        self._file = path
        self._saved_key = self._content_key()
        self._dismiss_resume()
        self._title_paint()
        self._status.set(f"Opened — {os.path.basename(path)}")

    def _dirty(self):
        if not self._autosave_armed:
            return False
        key = self._content_key()
        if self._saved_key is None:
            return any(getattr(self, v).get().strip()
                       for v in ("client_name", "opp_id", "project_name"))
        return key != self._saved_key

    def _confirm_discard(self):
        if not self._dirty():
            return True
        ans = messagebox.askyesnocancel(
            "Unsaved changes",
            "Save the current offer before going on?")
        if ans is None:
            return False
        if ans:
            self._save_json()
            return not self._dirty()
        return True

    def _title_paint(self, dirty=None):
        dirty = self._dirty() if dirty is None else dirty
        name = os.path.basename(self._file) if self._file else "New offer"
        self.root.title(f"{name}{'  •' if dirty else ''}  —  {APP_NAME} · "
                        f"Baker Hughes Thermodyn · DTI")

    def _apply(self, d):
        def s(var,key):
            if key in d and d[key] is not None: var.set(d[key])
        self._code_guard = True
        self.code_manual.set(bool(d.get("code_manual")))
        self._code_guard = False
        for var,key in [(self.offer_category,"offer_category"),(self.opp_id,"opp_id"),
                        (self.revision,"revision"),(self.gen_comment,"comment"),
                        (self.application,"application"),(self.gas_type,"gas_type"),
                        (self.custom_gas,"custom_gas"),(self.client_name,"client_name"),
                        (self.client_loc,"client_loc"),(self.project_date,"project_date"),
                        (self.ae_name,"ae_name"),(self.ae_phone,"ae_phone"),
                        (self.ae_email,"ae_email"),(self.sales_name,"sales_name"),
                        (self.sales_phone,"sales_phone"),(self.sales_email,"sales_email"),
                        (self.adv_name,"adv_name"),(self.adv_phone,"adv_phone"),
                        (self.adv_email,"adv_email"),
                        (self.sales_title,"sales_title"),(self.adv_title,"adv_title"),
                        (self.ae_title,"ae_title"),
                        (self.nda_signed,"nda_signed"),(self.num_impellers,"num_impellers"),
                        (self.casing_size,"casing_size"),(self.num_phases,"num_phases"),
                        (self.motor_type,"motor_type"),
                        (self.back_to_back,"back_to_back"),(self.rating,"rating"),
                        (self.ucp_type,"ucp_type"),(self.plc_option,"plc_option"),
                        (self.plc_safety,"plc_safety"),(self.thrust_bearing,"thrust_bearing"),
                        (self.motor_cooling_tcv,"motor_cooling_tcv"),
                        (self.motor_cooling_fan,"motor_cooling_fan"),
                        (self.vfd_cooling,"vfd_cooling"),(self.transfo_type,"transfo_type"),
                        (self.power,"power"),(self.cooling_system,"cooling_system"),
                        (self.antisurge_valve,"antisurge_valve"),
                        (self.filter_type,"filter_type"),(self.ehouse,"ehouse"),
                        (self.mcc,"mcc"),(self.comp_enclosure,"comp_enclosure"),
                        (self.cost_base,"cost_base"),(self.margin_pct,"margin_pct"),
                        (self.warr_pct,"warr_pct"),(self.lds_pct,"lds_pct"),
                        (self.project_name,"project_name"),
                        (self.nb_units,"nb_units"),(self.pdf_frame,"pdf_frame"),
                        (self.skeleton_file,"skeleton_file"),
                        (self.pdf_output,"pdf_output"),
                        (self.pdf_use_word,"pdf_use_word"),
                        (self.incoterm,"incoterm"),
                        (self.lead_time,"lead_time"),
                        (self.user_root,"user_root"),(self.db_root,"db_root"),
                        (self.use_db,"use_db"),(self.auto_place,"auto_place")]:
            s(var,key)
        for code,val in (d.get("doc_use_db") or {}).items():
            self._doc_var(code).set(bool(val))
        if "clarifications" in d:
            self._clarif.delete("1.0",tk.END); self._clarif.insert("1.0",d["clarifications"])
        self.lib_state = {k: dict(v) for k, v in (d.get("lib_state") or {}).items()}
        self.lib_choice = {k: v for k, v in (d.get("lib_choice") or {}).items()
                           if v and os.path.exists(v)}
        self._project_pick = d.get("project_pick") or ""
        if d.get("cooling_system") and d.get("cooling_system") not in COOLING_CONCEPTS:
            self.cooling_system.set("")
            self._status.set(f"Cooling concept “{d['cooling_system']}” no longer "
                             "exists: choose one of the seven concepts.")
        if self.incoterm.get() not in INCOTERMS:
            old = self.incoterm.get()
            self.incoterm.set("FCA")
            self._status.set(f"{old} is no longer offered: Incoterm set to FCA.")
        if re.fullmatch(r"\d{2}", self.casing_size.get().strip()):
            self.casing_size.set(self.casing_size.get().strip() + "0")
        self._lib_cat = None
        self.doc_attach = {k: list(v) for k, v in
                           (d.get("doc_attach") or {}).items()}
        self.doc_layout = dict(d.get("doc_layout") or {})
        self._struct_cache = (None, None)
        # one motor cooling: a file saved with both ticked reads as TCV
        self.motor_cooling.set(
            d.get("motor_cooling") if d.get("motor_cooling") in COOLING_TEXT
            else "Fan" if d.get("motor_cooling_fan") and not d.get("motor_cooling_tcv")
            else "TCV")
        self._auto_code()
        if self.code_manual.get() and d.get("config_code"):
            self._code_guard = True
            self.config_code.set(d["config_code"])
            self._code_guard = False
        self._code_hint_paint()
        self._update_nda(); self._on_category(); self._on_plc()
        self._scope_sync()
        self._scope_restore(d.get("scope_state"), d.get("scope_extra"))
        mode = d.get("price_mode")
        if mode == "manual" and d.get("final_price"):
            self._price_busy = True
            self.final_price.set(d["final_price"])
            self._price_busy = False
        self._set_price_mode(mode)
        self._paint_output(); self._paint_skeleton()
        self._scan_bg(silent=True)

    def _reset(self):
        """Start a new offer. The current one is kept if the user wants."""
        if not self._confirm_discard():
            return
        try:
            os.remove(self._autosave_path())
        except OSError:
            pass
        if self._tick_id:
            try:
                self.root.after_cancel(self._tick_id)
            except Exception:
                pass
        root = self.root
        for w in root.winfo_children():
            w.destroy()
        self.__dict__.clear()           # nothing of the old offer survives
        self.__init__(root)


# ─────────────────────────────────────────────────────────────────────────────
def main():
    # Crisp text on high-DPI Windows displays instead of a blurry upscale.
    try:
        if sys.platform.startswith("win"):
            from ctypes import windll
            try:
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                windll.user32.SetProcessDPIAware()
    except Exception:
        pass

    root = tk.Tk()
    root.withdraw()
    if os.environ.get("YBTOOL_NO_MOTION"):
        UI["motion"] = False
    _resolve_fonts(root)
    try:
        root.option_add("*Font", FONT_N)
    except Exception:
        pass

    BudgetOfferApp(root)

    # Open centred on the active screen, never larger than it.
    root.update_idletasks()
    w = min(int(1700 * UI["scale"]), root.winfo_screenwidth() - 40)
    h = min(int(960 * UI["scale"]), root.winfo_screenheight() - 80)
    x = (root.winfo_screenwidth() - w) // 2
    y = max((root.winfo_screenheight() - h) // 2 - 20, 0)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.deiconify()
    root.mainloop()

if __name__=="__main__":
    main()
