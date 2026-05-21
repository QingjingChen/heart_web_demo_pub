# HEART Benchmark Auditor — Web Demo

A single-file HTML user interface for the **HEART** (Human-centric Ethical
Assessment and Revision Toolkit) benchmark auditor: upload a benchmark or
dataset, score it on the 14-rubric scheme, see which tool families repair the
weak rubrics, and inspect generated adaptation examples per label.

> **Status.** This repository ships the **frontend only**. All buttons in the
> UI POST to a local backend at `http://127.0.0.1:8765` (see the `apiBase()`
> function near the bottom of `index.html`). Without that backend running, the
> page renders but the audit / adaptation actions will return errors.

## Quick look

Just open `index.html` in any modern browser — no build step, no dependencies,
no CDN calls. You will see the full UI; live audits require the companion
backend (not included in this repo).

## Live preview via GitHub Pages

Because the file is named `index.html` at the repo root, enabling GitHub Pages
on the `main` branch will publish the UI at
`https://<user>.github.io/<repo>/`. Note the same backend caveat applies: the
hosted page can render but cannot complete an audit unless the visitor is
running the backend locally.

## Endpoints the UI expects

If you build your own backend, the page calls:

| Method | Path | Purpose |
|--------|------|---------|
| `POST` | `/api/llm-test`       | Validate the entered LLM API key |
| `POST` | `/api/fusion-example` | Run the 14-rubric audit on an uploaded benchmark / dataset and return matched tool families |
| `POST` | `/api/adapt-sample`   | Generate domain-specific adaptation examples for a selected sample and selected tool patterns |

Payload shapes can be read directly from `index.html` (the JSON bodies are
constructed inline in the `runAudit`, `testApiKey`, and
`generateSelectedAdaptations` functions).

## What this demo is *not*

- It is **not** a self-contained tool — it has no offline scoring logic; all
  rubric scores, tool recommendations, and adaptation examples come from the
  backend.
- It is **not** an evaluation framework — the underlying methodology is
  described in the accompanying paper (under review).
- It does **not** transmit API keys to any third party. The key entered in the
  UI is sent only to the configured local backend (`127.0.0.1:8765` when
  opened via `file://`, or same-origin when served from a domain).

## License

MIT — see [LICENSE](LICENSE).

## Citation

A paper describing the HEART methodology, rubric scheme, and toolkit is under
review at a peer-reviewed venue. Citation details will be added here after
publication.
