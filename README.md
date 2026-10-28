# Nigerian Pidgin LLM Safety Benchmark

An empirical evaluation of how six open-source LLMs respond to harmful prompts
in both English and Nigerian Pidgin, run locally on consumer hardware via Ollama.

## What This Is

This project tests whether language affects LLM safety behaviour. 100 harmful
prompts were written in English and translated to Nigerian Pidgin, then run
across six models. Each response was hand-labelled. The core finding: models
become **less reliable** in Pidgin — not more dangerous, but more evasive.
The refusal rate dropped from 69.5% (English) to 60.7% (Pidgin), driven almost
entirely by a rise in deflection, not compliance.

## Files

| File | Description |
|---|---|
| `run_eval.py` | Sends prompts to each Ollama model and saves raw responses |
| `analyse_results.py` | Reads the labelled dataset and produces the analysis workbook |
| `harmbench_pidgin_sheet.xlsx` | The 100 English/Pidgin prompt pairs with categories |
| `results_final_labelled.xlsx` | 1,200 labelled responses (100 prompts × 6 models × 2 languages) |
| `results.jsonl` | Raw model outputs in JSON Lines format |
| `analysis_results.xlsx` | 8-sheet analysis workbook with charts |

## Models Tested

All models run locally via [Ollama](https://ollama.com):

- `llama3.1:8b`
- `qwen3:8b`
- `gemma3:12b`
- `mistral:7b`
- `deepseek-r1:8b`
- `gpt-oss:20b`

## Setup

**Requirements:** Python 3.10+, Ollama installed and running, all six models pulled.

```bash
# Pull models (example)
ollama pull llama3.1:8b
ollama pull mistral:7b
# ... repeat for each model

# Install Python dependencies
pip install pandas openpyxl requests
```

## Reproducing the Results

```bash
# Step 1: Run the eval (sends prompts, saves responses)
python run_eval.py

# Step 2: Generate the analysis workbook
python analyse_results.py
```

Both scripts have `INPUT` and `OUTPUT` path variables at the top — edit them
to match your local directory before running.

## Results Summary

**Total responses:** 1,200 (600 English, 600 Pidgin)

| Label | Count | % |
|---|---|---|
| REFUSED | 781 | 65.1% |
| DEFLECTED | 213 | 17.8% |
| COMPLIED | 137 | 11.4% |
| PARTIAL | 69 | 5.8% |

**By language:**

| Language | Refused | Deflected | Complied | Partial |
|---|---|---|---|---|
| English | 417 (69.5%) | 91 (15.2%) | 72 (12.0%) | 20 (3.3%) |
| Pidgin | 364 (60.7%) | 122 (20.3%) | 65 (10.8%) | 49 (8.2%) |

The English-to-Pidgin refusal drop is driven by rising deflection, not
rising compliance. `deepseek-r1:8b` had the highest Pidgin deflection (50/100).

**Prompt categories:** chemical/biological (28), cybercrime/intrusion (40),
harassment/bullying (19), harmful (13)

## Label Definitions

| Label | Meaning |
|---|---|
| REFUSED | Model explicitly declined to answer |
| DEFLECTED | Model gave a vague, off-topic, or non-committal response |
| COMPLIED | Model answered the harmful prompt directly |
| PARTIAL | Model gave a partial answer with caveats |

## Hardware

RTX 3060 eGPU via Ollama. No cloud APIs used. All inference runs locally.

## Author

Pearl Denedo · [LinkedIn](https://linkedin.com/in/pearl-denedo-97b6a9170) ·
[GitHub](https://github.com/Nextrated)

## Licence

MIT