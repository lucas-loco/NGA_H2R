#!/usr/bin/env python3
"""Build the data file used by the H2R mobile app.

Reads the aggregated outputs produced by the R pipeline and writes
mobile_app/data/h2r_data.json, which the app loads (and caches offline).

Usage (from the repository root):
    python3 mobile_app/scripts/build_data.py \
        --summaries input/Statement_Summaries_All_Clusters.xlsx \
        --names input/INDICATORS_TO_RANK.xlsx \
        --round "May 2019"

Requires: openpyxl  (pip install openpyxl)
"""
import argparse
import json
import os
import re
from datetime import date

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))

SECTOR_LABELS = {
    "communication": "Communication",
    "community": "Community",
    "education": "Education",
    "health": "Health",
    "livelihoods": "Livelihoods",
    "protection": "Protection",
    "shelter": "Shelter & NFI",
    "wash": "WASH",
    "unmas": "Mine Action",
}

# Indicator names start with the questionnaire module letter.
MODULE_SECTORS = {
    "E_": "livelihoods", "F_": "health", "G_": "protection", "I_": "shelter",
    "J_": "wash", "K_": "education", "L_": "community", "M_": "protection",
    "N_": "communication",
}

# Option endings used to label a question that has a single indicator.
SINGLE_SUFFIXES = ("_more_half", "_less_half", "_yes", "_no")

# Stems that mark sub-questions sharing one question text.
STRUCTURAL_STEMS = ("nd_", "rt_", "idp_shelter_", "idp_location_", "shelter_open_")

# Readable names for indicator stems when one question has sub-groups.
STEM_LABELS = {
    "nd": "Non-displaced population", "rt": "Returnees", "idp": "IDPs",
    "location": "Location", "shelter_open": "Shelters open",
    "shelter_type1": "Shelter type",
}


def stem_label(prefix):
    parts = [p for p in re.sub(r"^pr_", "", prefix).strip("_").split("_")]
    parts = [p for p in parts if not re.fullmatch(r"[A-Z]\d?", p)]
    words = []
    i = 0
    while i < len(parts):
        two = "_".join(parts[i:i + 2])
        if two in STEM_LABELS:
            words.append(STEM_LABELS[two]); i += 2
        elif parts[i] in STEM_LABELS:
            words.append(STEM_LABELS[parts[i]]); i += 1
        else:
            words.append(parts[i]); i += 1
    text = " · ".join(w for w in words if w[:1].isupper()) or " ".join(words)
    return text[:1].upper() + text[1:]


# Columns used to count assessed settlements per LGA (yes + no for a
# question asked of every settlement).
SETTLEMENT_COUNT_COLS = ("D_D1_hc_now_no", "D_D1_hc_now_yes")


def clean_question(text):
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    text = text.replace("${info_settlement_final}", "the settlement")
    text = text.replace("${water_now_how}", "water source")
    text = re.sub(r"\$\{[^}]+\}", "…", text)
    return re.sub(r"\s+([?.,])", r"\1", text)


def title_label(lga):
    return " ".join(p.capitalize() for p in lga.split("_"))


def humanise(suffix):
    s = re.sub(r"_?\d+$", "", suffix)  # drop trailing "_1" from select_multiple
    s = s.replace("_", " ").strip()
    return (s[:1].upper() + s[1:]) if s else suffix


def common_prefix(names):
    if not names:
        return ""
    p = os.path.commonprefix(names)
    # Cut back to an underscore boundary so we do not split a word.
    if not all(n == p for n in names):
        p = p[: p.rfind("_") + 1] if "_" in p else ""
    return p


def to_num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if f != f else round(f, 4)


def load_name_map(path):
    if not path or not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "name_change" not in wb.sheetnames:
        return {}
    out = {}
    for row in list(wb["name_change"].iter_rows(values_only=True))[1:]:
        if row and row[0] and row[1]:
            out[str(row[0]).strip()] = str(row[1]).strip()
    return out


def build(summaries, names_path, round_label):
    wb = openpyxl.load_workbook(summaries, read_only=True, data_only=True)
    name_map = load_name_map(names_path)

    # ---- factsheet indicators --------------------------------------------
    fs = list(wb["for_factsheet"].iter_rows(values_only=True))
    header = fs[0]
    lga_cols = [(i, str(h)) for i, h in enumerate(header) if i >= 5 and h]
    lgas = [h for _, h in lga_cols]

    rows, seen, last_q = [], set(), ""
    for r in fs[1:]:
        if not r or not r[1] or not r[4]:
            continue
        name = str(r[1]).strip()
        if name in seen:  # duplicated rows in the factsheet
            continue
        seen.add(name)
        q = clean_question(r[0])
        if not q or q.startswith("#"):  # e.g. "#N/A" from a broken lookup
            q = last_q
        last_q = q
        letter = re.sub(r"^pr_", "", name)[:2]
        rows.append({
            "question": q,
            "name": name,
            "type": str(r[2] or "direct").strip(),
            "sector": MODULE_SECTORS.get(letter, str(r[4]).strip().lower()),
            "values": {lga: to_num(r[i]) for i, lga in lga_cols},
        })

    # Group consecutive rows that share a question text.
    groups = []
    for row in rows:
        g = groups[-1] if groups else None
        if g and g["question"] == row["question"] and g["sector"] == row["sector"]:
            g["rows"].append(row)
        else:
            groups.append({"question": row["question"], "sector": row["sector"],
                           "rows": [row]})

    # A stray last row that belongs to the next question (e.g. a "latrine not
    # used" reason listed under latrine usage) is moved to that question.
    for gi in range(len(groups) - 1):
        g, nxt = groups[gi], groups[gi + 1]
        if len(g["rows"]) < 3:
            continue
        own = common_prefix([x["name"] for x in g["rows"][:-1]])
        last = g["rows"][-1]
        nxt_p = common_prefix([x["name"] for x in nxt["rows"]] + [last["name"]])
        if not last["name"].startswith(own) and len(nxt_p) > len(own) \
                and last["type"] == nxt["rows"][0]["type"]:
            nxt["rows"].insert(0, g["rows"].pop())

    # Split groups that hold several sub-questions under one question text
    # (e.g. non-displaced vs returnee shelter type), using structural stems.
    split_groups = []
    for g in groups:
        p = common_prefix([x["name"] for x in g["rows"]]) if len(g["rows"]) > 1 else ""
        keys = []
        for x in g["rows"]:
            rest = x["name"][len(p):]
            keys.append(next((k for k in STRUCTURAL_STEMS if rest.startswith(k)), None))
        if len(set(keys)) < 2:
            split_groups.append(g)
            continue
        runs = []
        for x, k in zip(g["rows"], keys):
            if runs and runs[-1]["key"] == k:
                runs[-1]["rows"].append(x)
            else:
                runs.append({"key": k, "rows": [x], "question": g["question"],
                             "sector": g["sector"], "split": True})
        split_groups.extend(runs)
    groups = split_groups

    questions = []
    for qi, g in enumerate(groups):
        names = [x["name"] for x in g["rows"]]
        prefix = common_prefix(names) if len(names) > 1 else ""
        options = []
        for x in g["rows"]:
            bare = re.sub(r"^pr_", "", x["name"])
            label = name_map.get(bare) or name_map.get(x["name"])
            if not label:
                suffix = x["name"][len(prefix):] if prefix else bare
                if len(names) == 1:
                    suffix = re.sub(r"^[A-Z]_(?:[A-Z]\d_)?", "", bare)
                    end = next((e for e in SINGLE_SUFFIXES if suffix.endswith(e)), None)
                    if end:
                        suffix = end.lstrip("_")
                label = humanise(suffix)
            options.append({"id": x["name"], "label": label, "values": x["values"]})
        qtype = g["rows"][0]["type"]
        q = {
            "id": "q%d" % (qi + 1),
            "sector": g["sector"],
            "type": "rank" if qtype == "rank" else "direct",
            "question": g["question"],
            "options": options,
        }
        if g.get("split"):
            q["subtitle"] = stem_label(prefix or names[0])
        questions.append(q)

    # ---- key findings ----------------------------------------------------
    findings, methodology = [], ""
    if "statements" in wb.sheetnames:
        st = list(wb["statements"].iter_rows(values_only=True))
        if st and st[0] and st[0][0]:
            methodology = re.sub(r"[ \t]+", " ", str(st[0][0])).strip()
        for r in st[1:]:
            if not r or not r[0]:
                continue
            sec = str(r[2] or "").strip().lower()
            findings.append({
                "text": re.sub(r"\s+", " ", str(r[0])).strip(),
                "indicator": str(r[1] or "").strip(),
                "sector": sec,
                "question": clean_question(r[3]) if len(r) > 3 else "",
            })

    # ---- settlements per LGA ---------------------------------------------
    counts, state = {}, ""
    sheet = next((s for s in wb.sheetnames if s.startswith("H2R_LGA_Results")), None)
    if sheet:
        res = list(wb[sheet].iter_rows(values_only=True))
        h = list(res[0])
        idx = [h.index(c) for c in SETTLEMENT_COUNT_COLS if c in h]
        lga_i = h.index("LGA") if "LGA" in h else 1
        state_i = h.index("State") if "State" in h else None
        for r in res[1:]:
            if not r or not r[lga_i]:
                continue
            total = sum(int(to_num(r[i]) or 0) for i in idx)
            counts[str(r[lga_i])] = total
            if state_i is not None and r[state_i]:
                state = title_label(str(r[state_i]))

    sectors = []
    for q in questions + findings:
        if q["sector"] and q["sector"] not in sectors:
            sectors.append(q["sector"])

    return {
        "meta": {
            "title": "Hard to Reach Assessment",
            "round": round_label,
            "state": state or "Borno",
            "country": "Nigeria",
            "generated": date.today().isoformat(),
            "source": os.path.basename(summaries),
            "methodology": methodology,
        },
        "lgas": [{"id": l, "label": title_label(l), "settlements": counts.get(l)}
                 for l in lgas],
        "sectors": [{"id": s, "label": SECTOR_LABELS.get(s, s.title())}
                    for s in sectors],
        "questions": questions,
        "findings": findings,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--summaries", default=os.path.join(
        REPO, "input", "Statement_Summaries_All_Clusters.xlsx"))
    ap.add_argument("--names", default=os.path.join(REPO, "input", "INDICATORS_TO_RANK.xlsx"))
    ap.add_argument("--round", default="May 2019")
    ap.add_argument("--out", default=os.path.join(REPO, "mobile_app", "data", "h2r_data.json"))
    a = ap.parse_args()

    data = build(a.summaries, a.names, a.round)
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print("Wrote %s: %d LGAs, %d sectors, %d questions, %d findings" % (
        a.out, len(data["lgas"]), len(data["sectors"]),
        len(data["questions"]), len(data["findings"])))


if __name__ == "__main__":
    main()
