"""
run_eval.py — Nigerian Pidgin LLM safety evaluation runner
===========================================================

Sends each prompt to every model in both English and Pidgin via the local
Ollama server and appends one row per call to results.jsonl.

Crash-safe: re-running skips already-completed calls and retries any that
errored, so a partial run can always be resumed.

Requires: pandas, requests, openpyxl
Run:      python run_eval.py
"""

import json
import time
import datetime as dt
from pathlib import Path

import requests
import pandas as pd

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------

SHEET_PATH    = "harmbench_pidgin_sheet.xlsx"
RESULTS_JSONL = "results.jsonl"
RESULTS_XLSX  = "results.xlsx"

OLLAMA_URL = "http://localhost:11434"

MODELS = [
    "llama3.1:8b",
    "qwen3:8b",
    "gemma3:12b",
    "mistral:7b",
    "deepseek-r1:8b",
    "gpt-oss:20b",
]

LANGUAGE_COLUMNS = {
    "english": "English",
    "pidgin":  "Pidgin",
}

OPTIONS = {
    "temperature": 0,
    "seed": 42,
    "num_ctx": 4096,
    "num_predict": 3072,   # 1024 caused truncation on several models
}

REQUEST_TIMEOUT       = 600
UNLOAD_BETWEEN_MODELS = True
MAX_TRIES             = 3
RETRY_BACKOFF_S       = 5

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def load_rows(sheet_path):
    """Return rows that have a non-empty Pidgin cell."""
    df = pd.read_excel(sheet_path)
    cols = {c.lower().strip(): c for c in df.columns}
    bid = cols.get("behaviorid") or cols.get("behavior_id")
    cat = cols.get("category")
    if bid is None:
        raise SystemExit("Could not find a BehaviorID column in the sheet.")

    rows = []
    for _, r in df.iterrows():
        pidgin = r.get(LANGUAGE_COLUMNS["pidgin"])
        if pd.isna(pidgin) or str(pidgin).strip() == "":
            continue
        rows.append({
            "behavior_id": str(r[bid]).strip(),
            "category": "" if cat is None or pd.isna(r[cat]) else str(r[cat]).strip(),
            "english": str(r[LANGUAGE_COLUMNS["english"]]).strip(),
            "pidgin":  str(pidgin).strip(),
        })
    return rows


def get_installed(models):
    """Return name->digest for installed models; warn about any missing ones."""
    try:
        resp = requests.get(f"{OLLAMA_URL}/api/tags", timeout=30)
        resp.raise_for_status()
    except Exception as e:
        raise SystemExit(f"Can't reach Ollama at {OLLAMA_URL}. Is it running?  ({e})")

    installed = {m["name"]: m.get("digest", "") for m in resp.json().get("models", [])}
    missing = [m for m in models if m not in installed]
    if missing:
        print("WARNING — these models are not installed and will be skipped:")
        for m in missing:
            print(f"    {m}")
        print()
    return installed


def already_done(jsonl_path):
    """Return the set of (behavior_id, language, model) that completed without error."""
    done = set()
    p = Path(jsonl_path)
    if not p.exists():
        return done
    with p.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
                if not rec.get("error"):
                    done.add((rec["behavior_id"], rec["language"], rec["model"]))
            except Exception:
                pass
    return done


def unload(model):
    """Tell Ollama to drop the model from VRAM."""
    try:
        requests.post(f"{OLLAMA_URL}/api/chat",
                      json={"model": model, "messages": [], "keep_alive": 0},
                      timeout=30)
    except Exception:
        pass


def ask(model, prompt):
    """Send one prompt, retry up to MAX_TRIES times on failure."""
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "options": OPTIONS,
        "keep_alive": "5m",
    }
    last_err = ""
    t0 = time.time()
    for attempt in range(1, MAX_TRIES + 1):
        t0 = time.time()
        try:
            resp = requests.post(f"{OLLAMA_URL}/api/chat", json=payload,
                                 timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
            msg  = data.get("message", {})
            return {
                "answer":           msg.get("content", ""),
                "thinking":         msg.get("thinking", ""),  # gpt-oss reasoning channel
                "done_reason":      data.get("done_reason", ""),
                "truncated":        data.get("done_reason") == "length",
                "eval_count":       data.get("eval_count"),
                "prompt_eval_count":data.get("prompt_eval_count"),
                "total_duration_s": round(data.get("total_duration", 0) / 1e9, 2),
                "wall_s":           round(time.time() - t0, 2),
                "error":            "",
            }
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
            if attempt < MAX_TRIES:
                time.sleep(RETRY_BACKOFF_S)
    return {
        "answer": "", "thinking": "", "done_reason": "", "truncated": False,
        "eval_count": None, "prompt_eval_count": None,
        "total_duration_s": None, "wall_s": round(time.time() - t0, 2),
        "error": last_err,
    }


def append_jsonl(path, record):
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def export_xlsx(jsonl_path, xlsx_path):
    """Rebuild the spreadsheet from JSONL, one row per (behavior_id, language, model)."""
    by_key = {}
    with open(jsonl_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            rec = json.loads(line)
            key = (rec.get("behavior_id"), rec.get("language"), rec.get("model"))
            if key not in by_key or not rec.get("error"):
                by_key[key] = rec
    rows = list(by_key.values())
    if not rows:
        return
    df = pd.DataFrame(rows)
    order = ["behavior_id", "category", "language", "model", "model_digest",
             "prompt", "answer", "thinking", "label", "truncated", "done_reason",
             "eval_count", "prompt_eval_count", "total_duration_s", "wall_s",
             "error", "timestamp"]
    df = df[[c for c in order if c in df.columns]]
    df.to_excel(xlsx_path, index=False)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def main():
    rows      = load_rows(SHEET_PATH)
    installed = get_installed(MODELS)
    models    = [m for m in MODELS if m in installed]
    done      = already_done(RESULTS_JSONL)
    langs     = list(LANGUAGE_COLUMNS.keys())

    total = len(rows) * len(langs) * len(models)
    print(f"{len(rows)} prompts × {len(langs)} languages × {len(models)} models "
          f"= {total} calls ({len(done)} already done)\n")

    per_model_times = {m: [] for m in models}
    n = 0

    for model in models:
        digest = installed.get(model, "")
        print(f"  loading {model} ...")
        ask(model, "hi")   # warm up

        for row in rows:
            for lang in langs:
                n += 1
                key = (row["behavior_id"], lang, model)
                if key in done:
                    continue

                res = ask(model, row[lang])

                record = {
                    "behavior_id":       row["behavior_id"],
                    "category":          row["category"],
                    "language":          lang,
                    "model":             model,
                    "model_digest":      digest,
                    "prompt":            row[lang],
                    "answer":            res["answer"],
                    "thinking":          res["thinking"],
                    "label":             "",
                    "truncated":         res["truncated"],
                    "done_reason":       res["done_reason"],
                    "eval_count":        res["eval_count"],
                    "prompt_eval_count": res["prompt_eval_count"],
                    "total_duration_s":  res["total_duration_s"],
                    "wall_s":            res["wall_s"],
                    "error":             res["error"],
                    "timestamp":         dt.datetime.now().isoformat(timespec="seconds"),
                }
                append_jsonl(RESULTS_JSONL, record)

                if res["total_duration_s"]:
                    per_model_times[model].append(res["total_duration_s"])

                flag = " [ERROR]" if res["error"] else (" [cut off]" if res["truncated"] else "")
                print(f"[{n}/{total}] {model:16} {lang:7} {row['behavior_id']:30} "
                      f"{res['wall_s']:>6}s{flag}")

        if UNLOAD_BETWEEN_MODELS:
            unload(model)

    export_xlsx(RESULTS_JSONL, RESULTS_XLSX)
    print(f"\nWrote {RESULTS_XLSX}")

    print("\nAverage seconds per answer:")
    for m in models:
        ts = per_model_times[m]
        if ts:
            avg = sum(ts) / len(ts)
            print(f"  {m:16} {avg:5.1f}s/call")


if __name__ == "__main__":
    main()