"""Consolidate <OUTPUT_DIR>/results/*.json into summary.json and a Markdown table (summary.md)."""
import glob
import json
import os

from settings import PERSONA_LABELS, RESULTS_DIR, SITES


def main():
    rows, data = [], {}
    for path in sorted(glob.glob(os.path.join(RESULTS_DIR, "*.json"))):
        who = os.path.splitext(os.path.basename(path))[0]
        if who not in PERSONA_LABELS:
            continue
        with open(path, encoding="utf-8") as f:
            result = json.load(f)
        name, agents, perms = PERSONA_LABELS[who]
        data[who] = {}
        for letter in SITES:
            cell = result["sites"].get(letter)
            if not cell:
                continue
            if cell["access"] != "ok":
                out = {"page": "ACCESS DENIED", "A_sdk": "-", "B_sdk": "-", "B_embed": "-"}
            else:
                out = {"page": "OK", "A_sdk": cell["agentA_sdk"]["verdict"], "B_sdk": cell["agentB_sdk"]["verdict"],
                       "B_embed": cell["agentB_embed"]["verdict"],
                       "A_greeting": (cell["agentA_sdk"].get("greeting") or [""])[0]}
            data[who][letter] = out
            rows.append((name, agents, letter, perms[letter], out))
    lines = ["| Persona | Agent sharing | Site | Site permission | Page | Agent A via Agents SDK | Agent B via Agents SDK | Agent B via hosted embed |",
             "|---|---|---|---|---|---|---|---|"]
    lines += [f"| {n} | {a} | Site {s} | {p} | {o['page']} | {o['A_sdk']} | {o['B_sdk']} | {o['B_embed']} |" for n, a, s, p, o in rows]
    with open(os.path.join(RESULTS_DIR, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, ensure_ascii=False)
    with open(os.path.join(RESULTS_DIR, "summary.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
