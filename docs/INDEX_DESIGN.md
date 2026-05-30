# `heart_index.json` — design notes

This document explains how the HEART web demo combines four data layers
to produce **revised benchmark items with provenance**. The design
intentionally separates *what the LLM is allowed to do* from *what the
workbook says is correct*: the LLM extracts and fills templates; the
workbook index constrains, retrieves, and grades.

## The four data layers

```
┌───────────────────────────────────────────────────────────────────┐
│ 1. Uploaded paper / GitHub repo / dataset                         │
│    PDF text, CSV/JSON fields, README. Source of:                  │
│      benchmark profile, claim, task, metric, label schema.        │
└──────────────────────────┬────────────────────────────────────────┘
                           │ (LLM extraction)
                           ▼
┌───────────────────────────────────────────────────────────────────┐
│ 2. heart_index.json (this index)                                  │
│    Authoritative constraint + retrieval layer:                    │
│      14 rubrics + calibrated anchors + weak triggers              │
│      20 repair tools + output-field schemas                       │
│      indices.rubric_to_tools / dimension_to_tools                 │
│      diagnosis_rules (79 rules; LLM-free)                         │
│      6 quality metrics                                            │
└──────────────────────────┬────────────────────────────────────────┘
                           │ (rule-based retrieve)
                           ▼
┌───────────────────────────────────────────────────────────────────┐
│ 3. Target-domain template (Medical / Finance / Education)         │
│    Required fields + normative anchors + failure modes.           │
│    Constrains what a revised item must contain.                   │
└──────────────────────────┬────────────────────────────────────────┘
                           │ (template-fill)
                           ▼
┌───────────────────────────────────────────────────────────────────┐
│ 4. LLM draft (Qwen / OpenAI compatible)                           │
│    Produces revised items. Must:                                  │
│      - emit only fields named in step 3,                          │
│      - cite normative_anchor and failure_mode from the template,  │
│      - attach `source: heart_index.<...>` provenance,             │
│      - leave any decision the index does not authorise blank.     │
└───────────────────────────────────────────────────────────────────┘
```

## How a request flows through the backend

For an end-to-end "upload paper → revised cases" request the demo
performs roughly the following steps. Existing `/api/analyze`,
`/api/fusion-example`, and `/api/adapt-sample` endpoints are unchanged
in shape — they now additionally consult the index and embed
`source` / `tool_id` / `index_evidence` fields in their replies.

1. **Layer 1 — extraction.** `parse_benchmark_bytes` reads the upload
   (PDF, CSV, JSONL, XLSX, ZIP) into a benchmark profile (claim, task,
   metric, label schema, dataset description).
2. **Layer 2 — rule-based diagnosis.** `heart_index_loader.match_triggers`
   greps the profile text against `indices.trigger_to_rubric`. Each hit
   produces a `(rubric_id, matches_score, rationale)` finding whose
   provenance points at `heart_index.diagnosis_rules.<rule_id>`. No LLM
   call is required for this step.
3. **Layer 2 — tool retrieval.** For each weak rubric,
   `recommend_tools_for_rubric` returns the canonical tools listed in
   `indices.rubric_to_tools` together with the `output_fields` schema
   that every revised item should carry. Provenance points at
   `heart_index.tools.<tool_id>`.
4. **Layer 3 — domain template.** The user selects one of the published
   `domain_templates` (currently `medical`, `finance`, `education`).
   The template supplies the *required_fields*, the
   *normative_anchors*, the *failure_modes*, and the *action_labels*
   that the revised items must conform to.
5. **Layer 4 — LLM fill.** The LLM is given the original benchmark
   item, the chosen tool's `output_fields`, and the domain template,
   and is asked to emit a *list of revised items*. The prompt
   constrains the LLM to:
   - emit only fields named in `output_fields ∪ required_fields`,
   - cite a `normative_anchor.id` from the domain template,
   - tag at least one `failure_mode.id` from the domain template,
   - copy through the `source` pointer.
6. **Coverage scoring.** `heart_index_loader.compute_metrics` is run on
   the resulting list to produce the 6 LLM-free quality metrics
   defined in `index.metrics`.

## Provenance contract

Every record the backend emits that draws on the index MUST carry a
`source` field shaped like one of:

```
heart_index.rubrics.<rubric_id>
heart_index.tools.<tool_id>
heart_index.diagnosis_rules.<rule_id>
heart_index.domain_templates.<domain_id>
```

The front-end can deep-link the user back to the corresponding row in
the public Heart repository's `workbook/exports/`.

## API surface added by this change

| Method | Path | Purpose |
|--------|------|---------|
| `GET`  | `/api/index` | Return the full `heart_index.json`. The frontend can use it to populate dropdowns, show provenance pointers, and allow user override. |
| `GET`  | `/api/index/domain_templates` | Return only the `domain_templates` block (lighter payload for filling a dropdown). |
| `POST` | `/api/index/diagnose` | `{"text": "...", "domain_id": "medical"}` → rule-based weak-rubric findings + recommended tools. No LLM call. |
| `POST` | `/api/index/metrics` | `{"revised_items": [...], "domain_id": "medical"}` → 6 quality metrics. No LLM call. |

The existing endpoints (`/api/health`, `/api/analyze`,
`/api/fusion-example`, `/api/adapt-sample`) keep their response shape;
they just gain a `tool_id`, `output_fields`, `index_evidence`, and
`source` field on tools they recommend.

## Regenerating the index

`heart_index.json` is mirrored from the Heart resource repository
([`workbook/exports/heart_index.json`](https://github.com/QingjingChen/Heart/blob/main/workbook/exports/heart_index.json)).
The full regeneration procedure is documented at
[`workbook/exports/index_schema.md`](https://github.com/QingjingChen/Heart/blob/main/workbook/exports/index_schema.md)
in that repository.

## What this design is not

- **Not a benchmark grader.** The backend does not auto-score the
  user's benchmark. It diagnoses gaps and proposes repair tools whose
  recommendations are traceable; the grading is still a human task.
- **Not free-form generation.** The LLM cannot recommend a tool that
  is not in the index. If a profile matches no trigger, the response is
  an empty `findings` list, not an invented recommendation.
- **Not domain-complete.** Three domains are shipped today (medical,
  finance, education) because they have the strongest grounding in the
  104-benchmark adaptation examples. New domain templates can be added
  by following the structure in `index_schema.md` §`domain_templates`.
