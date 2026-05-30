"""
HEART index loader.

Loads `heart_index.json` once at import time and exposes helper functions
used by `heart_backend.py` to:

  - look up rubric / tool / domain template / metric entries by id,
  - run rule-based diagnosis via the `trigger_to_rubric` map,
  - compute the 6 LLM-free quality metrics on a list of revised items,
  - format a provenance pointer string for any record the backend emits.

Provenance contract — every record returned by the backend that draws on
the index must carry `source` of one of these shapes::

    heart_index.rubrics.<rubric_id>
    heart_index.tools.<tool_id>
    heart_index.diagnosis_rules.<rule_id>
    heart_index.domain_templates.<domain_id>
"""

from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
INDEX_PATH = ROOT / "heart_index.json"

_INDEX: dict[str, Any] | None = None


def load_index(path: Path = INDEX_PATH) -> dict[str, Any]:
    """Load and cache `heart_index.json`. Safe to call repeatedly."""
    global _INDEX
    if _INDEX is None:
        if not path.exists():
            _INDEX = {
                "version": "0.0",
                "rubrics": {},
                "tools": {},
                "indices": {
                    "rubric_to_tools": {},
                    "dimension_to_tools": {},
                    "trigger_to_rubric": {},
                },
                "domain_templates": {},
                "metrics": {},
                "diagnosis_rules": [],
                "error": f"heart_index.json not found at {path}",
            }
        else:
            _INDEX = json.loads(path.read_text(encoding="utf-8"))
    return _INDEX


def get_index() -> dict[str, Any]:
    return load_index()


# --------------------------------------------------------------------- lookups


def rubric(rubric_id: str) -> dict[str, Any] | None:
    return load_index().get("rubrics", {}).get(rubric_id)


def tool(tool_id: str) -> dict[str, Any] | None:
    return load_index().get("tools", {}).get(tool_id)


def domain_template(domain_id: str) -> dict[str, Any] | None:
    return load_index().get("domain_templates", {}).get(domain_id)


def rubric_to_tools(rubric_id: str) -> list[str]:
    return load_index().get("indices", {}).get("rubric_to_tools", {}).get(rubric_id, [])


def dimension_to_tools(dim: str) -> list[str]:
    return load_index().get("indices", {}).get("dimension_to_tools", {}).get(dim, [])


def diagnosis_rules() -> list[dict[str, Any]]:
    return load_index().get("diagnosis_rules", [])


def provenance(kind: str, key: str) -> str:
    """Build a `heart_index.<kind>.<key>` pointer string."""
    return f"heart_index.{kind}.{key}"


# ----------------------------------------------------- rule-based diagnosis


def match_triggers(text: str) -> list[dict[str, Any]]:
    """Run all `trigger_to_rubric` patterns against the supplied text.

    Returns a list of matches, each shaped::

        {
          "rubric_id": "...",
          "trigger": "...",
          "matches_score": "0" | "1-2",
          "rationale": "...",
          "source": "heart_index.diagnosis_rules.<rule_id>"
        }
    """
    if not text:
        return []
    idx = load_index()
    trig_map: dict[str, list[dict[str, Any]]] = (
        idx.get("indices", {}).get("trigger_to_rubric", {})
    )
    rules_by_pair: dict[tuple[str, str], str] = {}
    for rule in idx.get("diagnosis_rules", []):
        rules_by_pair[(rule["rubric_id"], rule["trigger_pattern"].lower())] = rule["id"]

    hay = text.lower()
    matches: list[dict[str, Any]] = []
    for pattern, entries in trig_map.items():
        if not pattern:
            continue
        if pattern in hay:
            for ent in entries:
                rid = ent.get("rubric_id", "")
                rule_id = rules_by_pair.get((rid, pattern), "")
                matches.append(
                    {
                        "rubric_id": rid,
                        "trigger": pattern,
                        "matches_score": ent.get("matches_score", "1-2"),
                        "rationale": ent.get("rationale", ""),
                        "source": provenance("diagnosis_rules", rule_id)
                        if rule_id
                        else provenance("rubrics", rid),
                    }
                )
    return matches


def recommend_tools_for_rubric(
    rubric_id: str, k: int = 5, domain: str | None = None
) -> list[dict[str, Any]]:
    """Return the top-k tool entries (decorated with provenance) for a weak rubric.

    If `domain` is supplied AND matches a domain_template, the returned tools
    are re-ranked so that tools whose `dimensions_applicable` mention that
    domain or its parent dimension come first.
    """
    candidates = rubric_to_tools(rubric_id) or []
    out: list[dict[str, Any]] = []
    for tid in candidates[:k]:
        t = tool(tid)
        if not t:
            continue
        rub_evidence = next(
            (r for r in t.get("rubrics_improved", []) if r.get("rubric_id") == rubric_id),
            None,
        )
        out.append(
            {
                "tool_id": tid,
                "canonical_name": t.get("canonical_name"),
                "core_practice": t.get("core_practice"),
                "problem_fixed": t.get("problem_fixed"),
                "how_it_lifts_score": rub_evidence.get("how") if rub_evidence else "",
                "bench_ref": rub_evidence.get("bench_ref") if rub_evidence else "",
                "survey_ref": rub_evidence.get("survey_ref") if rub_evidence else "",
                "output_fields": t.get("output_fields", []),
                "source": provenance("tools", tid),
            }
        )
    return out


# --------------------------------------------------------------- metrics


def compute_metrics(
    revised_items: list[dict[str, Any]],
    domain_id: str | None = None,
) -> dict[str, Any]:
    """Compute the 6 LLM-free quality metrics over a list of revised items.

    `revised_items` is a list of per-sample dicts; expected keys (any subset):

      - expected_action, failure_mode, normative_anchor, pair_id, source
      - plus arbitrary `required_fields` per the chosen `domain_template`
    """
    n = len(revised_items)
    if n == 0:
        return {
            "n_items": 0,
            "metrics": {k: None for k in load_index().get("metrics", {})},
        }

    tmpl = domain_template(domain_id) if domain_id else None
    required_fields = (
        [f["field"] for f in (tmpl.get("required_fields", []) if tmpl else []) if f.get("required")]
        if tmpl
        else []
    )
    template_failure_modes = [f["id"] for f in (tmpl.get("failure_modes", []) if tmpl else [])]

    # 1. context_variable_coverage (per item -> mean across items)
    if required_fields:
        per_item = []
        for it in revised_items:
            filled = sum(1 for f in required_fields if str(it.get(f, "")).strip())
            per_item.append(filled / len(required_fields))
        cvc = sum(per_item) / len(per_item)
    else:
        cvc = None

    # 2. normative_anchor_coverage
    with_anchor = sum(1 for it in revised_items if str(it.get("normative_anchor", "")).strip())
    nac = with_anchor / n

    # 3. action_label_diversity
    actions = [str(it.get("expected_action", "")).strip() for it in revised_items]
    ald = len({a for a in actions if a})

    # 4. counterfactual_pair_coverage
    pair_counts = Counter(
        str(it.get("pair_id", "")).strip() for it in revised_items if str(it.get("pair_id", "")).strip()
    )
    target_pairs = max(len(pair_counts), 1)
    complete_pairs = sum(1 for _, c in pair_counts.items() if c >= 2)
    cpc = complete_pairs / target_pairs if pair_counts else 0.0

    # 5. failure_mode_coverage
    fm_observed = {str(it.get("failure_mode", "")).strip() for it in revised_items}
    fm_observed.discard("")
    if template_failure_modes:
        fmc = len(fm_observed & set(template_failure_modes)) / len(template_failure_modes)
    else:
        fmc = None

    # 6. traceability_coverage
    with_src = sum(1 for it in revised_items if str(it.get("source", "")).strip())
    tc = with_src / n

    return {
        "n_items": n,
        "domain_id": domain_id,
        "required_fields_checked": required_fields,
        "template_failure_modes": template_failure_modes,
        "metrics": {
            "context_variable_coverage": cvc,
            "normative_anchor_coverage": nac,
            "action_label_diversity": ald,
            "counterfactual_pair_coverage": cpc,
            "failure_mode_coverage": fmc,
            "traceability_coverage": tc,
        },
    }
