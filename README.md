# HEART Benchmark Auditor — Web Demo

An interactive front-end for the **HEART** benchmark-revision workflow. This
demo is the runnable companion to the canonical resource repository:

> **Resource repository:** <https://github.com/QingjingChen/Heart>
> — 5 policy-grounded dimensions, 14 audit rubrics, a 6-tool repair workbox,
> and a meta-review of 103 benchmark / survey papers.

The demo takes a benchmark file you upload, audits it against the 14 rubrics,
recommends one or more of the 6 workbox repair tools, and — optionally with
your own LLM API key — drafts a concrete revision plan and per-sample
adaptation examples in the target domain you specify.

## What the demo implements

| Step | What happens | Resource section |
|---|---|---|
| 1 — Upload benchmark | Parse `JSONL / JSON / CSV / TSV / XLSX / ZIP / TXT / MD`, extract a profile (constructs, item types, label sources) | `docs/01_problem.md` |
| 2 — Pick target labels | Multi-select from the 5 dimensions: Human-Centered / Fairness & Inclusiveness / Safety & Reliability / Trustworthiness & Controllability / Privacy Protection | `docs/02_five_dimensions.md` |
| 3 — Rubric diagnosis | Score the benchmark against each of the 14 rubrics (3-layer: content validity, evaluation design, governance reliability) | `docs/03_fourteen_rubrics.md` |
| 4 — Tool recommendation | Match the weak rubrics to the 6 repair tools (Claim–Construct Mapper, Normative Source Ledger, Scenario & Stakeholder Expander, Disagreement & Ground-Truth Protocol, Metric & Failure-Mode Reporter, Lifecycle & Contamination Guard) | `docs/04_workbox_six_tools.md` |
| 5 — Revision plan | Generate a per-sample adaptation example fused with your uploaded domain dataset; LLM-drafted if a key is provided | `docs/05_guidebook_workflow.md` |

The doc references point into the companion resource repository.

## Contents of this repo

| Path | Purpose |
|------|---------|
| `index.html` | Static browser UI — no build step, no CDN dependency |
| `heart_backend.py` | Local FastAPI server: parsing, rubric scoring, tool matching, adaptation generation |
| `workbooks/科技伦理toolkit.xlsx` | Backend-compatible toolkit workbook (mirrored from the resource repo) |
| `workbooks/HEART_Excel_含原始与修订.xlsx` | Full workbook deliverable with original and revised sheets |
| `scripts/` | Workbook-maintenance utilities (lineage extraction, per-rubric mapping) |
| `docs/API.md` | Local API examples |
| `requirements.txt` | Python dependencies |
| `CITATION.cff` | Citation metadata (anonymous; finalised at resource release) |

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn heart_backend:app --host 127.0.0.1 --port 8765
```

Then open <http://127.0.0.1:8765/> in a browser.

You can also open `index.html` directly. In that mode the page sends API
calls to `http://127.0.0.1:8765`, so the backend still needs to be running.

## Endpoints the UI calls

| Method | Path | Purpose |
|--------|------|---------|
| `POST` | `/api/llm-test`       | Validate the entered LLM API key |
| `POST` | `/api/fusion-example` | Run the 14-rubric audit on an uploaded benchmark / dataset and return matched tool families |
| `POST` | `/api/adapt-sample`   | Generate domain-specific adaptation examples for a selected sample and selected tool patterns |
| `POST` | `/api/analyze`        | Single-benchmark profile + 14-rubric diagnosis + tool recommendations + revision plan |

Full request / response examples are in [`docs/API.md`](docs/API.md).

## GitHub Pages

Because the UI is `index.html` at the repository root, GitHub Pages can render
the demo shell directly. Interactive audit actions still require the local
backend at `http://127.0.0.1:8765` (or another base, see below).

For a different API base, set it in the browser console:

```js
localStorage.setItem("heartApiBase", "http://127.0.0.1:8765")
```

## Secrets

Do not commit API keys. The backend accepts a key only in the local request
body and does not write it to disk. The working-folder `AccessKey.csv` and
local DashScope key files are deliberately excluded.

## License

MIT — see [LICENSE](LICENSE). The workbook mirrored into `workbooks/` is
licensed CC-BY-4.0 in the [resource repository](https://github.com/QingjingChen/Heart).

## Citation

A companion paper describing the HEART methodology, rubric scheme, and
workbox is in preparation. Final citation, venue, and DOI metadata will be
added to [`CITATION.cff`](CITATION.cff) and to the
[resource repository](https://github.com/QingjingChen/Heart)
at release time.
