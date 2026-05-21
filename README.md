# HEART Benchmark Auditor Demo

A compact public package for the **HEART** benchmark auditor demo. It includes
the static web UI, the local FastAPI backend, the final workbook deliverable,
and the final workbook-maintenance scripts retained from the working folder.

## Contents

| Path | Purpose |
|------|---------|
| `index.html` | Static browser UI; no build step and no CDN dependency. |
| `heart_backend.py` | Local API server for scoring, tool matching, and adaptation examples. |
| `workbooks/HEART_Excel_含原始与修订.xlsx` | Final, most complete workbook deliverable with original and revised sheets. |
| `workbooks/科技伦理toolkit.xlsx` | Backend-compatible toolkit workbook used by the API. |
| `scripts/` | Final lineage/per-rubric workbook scripts and their curated JSON support files. |
| `docs/API.md` | Local API examples. |

Old scoring scripts, backups, logs, local virtualenvs, access keys, raw model
outputs, paper zip files, and temporary LaTeX/package artifacts are intentionally
left out.

## Run Locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn heart_backend:app --host 127.0.0.1 --port 8765
```

Then open `http://127.0.0.1:8765/`.

You can also open `index.html` directly. In that mode, the page sends API calls
to `http://127.0.0.1:8765`, so the backend still needs to be running for live
audits.

## Endpoints the UI expects

| Method | Path | Purpose |
|--------|------|---------|
| `POST` | `/api/llm-test`       | Validate the entered LLM API key |
| `POST` | `/api/fusion-example` | Run the 14-rubric audit on an uploaded benchmark / dataset and return matched tool families |
| `POST` | `/api/adapt-sample`   | Generate domain-specific adaptation examples for a selected sample and selected tool patterns |

More examples are in [`docs/API.md`](docs/API.md).

## GitHub Pages

Because the UI is `index.html` at the repository root, GitHub Pages can render
the demo shell directly. Interactive audit actions still require a local backend
at `http://127.0.0.1:8765`.

For a different API base, set it in the browser console:

```js
localStorage.setItem("heartApiBase", "http://127.0.0.1:8765")
```

## Secrets

Do not commit API keys. The backend accepts a key only in the local request body
and does not write it to disk. The working-folder `AccessKey.csv` and local
DashScope key files are deliberately excluded.

## License

MIT — see [LICENSE](LICENSE).

## Citation

A paper describing the HEART methodology, rubric scheme, and toolkit is under
review at a peer-reviewed venue. Citation details will be added here after
publication.
