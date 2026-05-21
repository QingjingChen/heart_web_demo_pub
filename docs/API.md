# HEART Local API

Base URL:

```text
http://127.0.0.1:8765
```

## Analyze One Benchmark

```http
POST /api/analyze
Content-Type: application/octet-stream
X-Filename: benchmark.jsonl
X-Domain: Healthcare
X-Label: Auto detect
X-Policy: Internal ethics policy + EU AI Act
```

Returns benchmark profile, 14 HEART rubric diagnoses, recommended tools, and a revision plan.

## Analyze Benchmark And Dataset

```http
POST /api/fusion-example
Content-Type: application/json
```

Request body:

```json
{
  "domain": "Healthcare",
  "labels": ["Fairness & Inclusiveness", "Privacy Protection"],
  "policy": "Internal ethics policy + EU AI Act",
  "max_examples": 3,
  "llm": {
    "provider": "qwen-dashscope",
    "model": "qwen-plus",
    "api_key": "sk-...",
    "enabled": true
  },
  "benchmark": {
    "filename": "benchmark.jsonl",
    "content_base64": "..."
  },
  "dataset": {
    "filename": "domain_dataset.csv",
    "content_base64": "..."
  }
}
```

Use `"labels": []` or omit `labels` for automatic multi-label detection. The web UI sends multiple selected labels as this list.

Response includes the extracted single dataset sample in `datasetProfile.sample`, plus recommended tools. It does not generate adaptations yet; call `/api/adapt-sample` after selecting tools.

## Adapt One Extracted Sample

```http
POST /api/adapt-sample
Content-Type: application/json
```

Request body:

```json
{
  "domain": "Healthcare",
  "labels": ["Fairness & Inclusiveness"],
  "policy": "Internal ethics policy + EU AI Act",
  "sample": "text: one extracted original dataset sample",
  "tools": [
    {
      "title": "BBQ Context-Pair Adapter",
      "source": "BBQ (2022)",
      "toolKey": "BBQ Context-Pair Adapter::BBQ (2022)"
    }
  ],
  "profile": {},
  "datasetProfile": {},
  "rubrics": [],
  "llm": {
    "provider": "qwen-dashscope",
    "model": "qwen-plus",
    "api_key": "sk-...",
    "enabled": true
  }
}
```

Response includes:

```json
{
  "sample": "text: one extracted original dataset sample",
  "llm": {
    "used": true,
    "message": "LLM-generated adaptations for the extracted sample"
  },
  "fusionExamples": [
    {
      "toolTitle": "BBQ Context-Pair Adapter",
      "source": "BBQ (2022)",
      "matchedRubrics": [],
      "inputDatasetSample": "...",
      "benchmarkPattern": {},
      "rewrittenItem": {},
      "fusionRationale": "...",
      "qualityChecks": []
    }
  ]
}
```

Python example:

```python
import base64
import json
import requests


def as_upload(path):
    with open(path, "rb") as f:
        return {
            "filename": path.split("/")[-1],
            "content_base64": base64.b64encode(f.read()).decode("ascii"),
        }


audit_payload = {
    "domain": "Healthcare",
    "labels": ["Fairness & Inclusiveness"],
    "policy": "Internal ethics policy + EU AI Act",
    "max_examples": 3,
    "benchmark": as_upload("benchmark.jsonl"),
    "dataset": as_upload("clinic_cases.csv"),
}

audit = requests.post("http://127.0.0.1:8765/api/fusion-example", json=audit_payload, timeout=60)
audit.raise_for_status()
audit_data = audit.json()

adapt_payload = {
    "domain": "Healthcare",
    "labels": audit_data["profile"].get("labelMatches", []),
    "sample": audit_data["datasetProfile"]["sample"],
    "tools": audit_data["tools"][:2],
    "profile": audit_data["profile"],
    "datasetProfile": audit_data["datasetProfile"],
    "rubrics": audit_data["rubrics"],
    "llm": {
        "provider": "qwen-dashscope",
        "model": "qwen-plus",
        "api_key": "YOUR_DASHSCOPE_KEY",
        "enabled": True,
    }
}

resp = requests.post("http://127.0.0.1:8765/api/adapt-sample", json=adapt_payload, timeout=60)
resp.raise_for_status()
print(json.dumps(resp.json()["fusionExamples"], ensure_ascii=False, indent=2))
```

Supported file types: `JSONL`, `JSON`, `CSV`, `TSV`, `XLSX`, `ZIP`, `TXT`, `MD`.

## Test Qwen/DashScope API Key

```http
POST /api/llm-test
Content-Type: application/json
```

```json
{
  "provider": "qwen-dashscope",
  "model": "qwen-plus",
  "api_key": "sk-..."
}
```

The key is used only for the current local request and is not written to disk by `heart_backend.py`.
