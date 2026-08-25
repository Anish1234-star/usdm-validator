
import os
import sys
import json

from google import genai
from google.genai import types

client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))


ELIGIBILITY_SCHEMA = {
    "type": "object",
    "properties": {
        "inclusionCriteria": {"type": "array", "items": {"type": "string"}},
        "exclusionCriteria": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["inclusionCriteria", "exclusionCriteria"],
}

ELIGIBILITY_PROMPT = """Find the Inclusion Criteria and Exclusion Criteria section(s) of this
clinical trial protocol (often titled "Eligibility Criteria," "Study Population," "Inclusion
and Exclusion Criteria," or similar).

Extract:
- "inclusionCriteria": each individual inclusion criterion as its own string, exactly as
  written (numbering like "1." or "IN01" can be dropped, but keep the actual criterion text
  complete and unabridged).
- "exclusionCriteria": same, for exclusion criteria.

Do not summarize or merge criteria together -- one array entry per distinct numbered/bulleted
criterion. If a criterion has sub-parts (a, b, c) that are only meaningful together, keep them
as one entry. If no such section exists in this document, return empty lists for both.

Respond with JSON only, matching the schema."""


ARMS_SCHEMA = {
    "type": "object",
    "properties": {
        "arms": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "type": {"type": "string"},
                    "description": {"type": "string"},
                    "interventions": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {"type": "string"},
                                "type": {"type": "string"},
                                "dose": {"type": "string"},
                                "route": {"type": "string"},
                                "frequency": {"type": "string"},
                            },
                            "required": ["name"],
                        },
                    },
                },
                "required": ["name", "description"],
            },
        },
    },
    "required": ["arms"],
}

ARMS_PROMPT = """Find the study arms / treatment groups section of this clinical trial protocol
(often in "Study Design," "Trial Design," or "Interventions" sections).

For each distinct arm/treatment group, extract:
- "name": the arm's name/label exactly as written (e.g. "Arm A", "Placebo Group",
  "Experimental: 200mg QD").
- "type": the arm's role if stated or clearly implied (e.g. "Experimental", "Active
  Comparator", "Placebo Comparator", "Sham Comparator"). Leave as an empty string if genuinely
  unclear.
- "description": a short description of what happens to subjects in this arm, as written in
  the protocol.
- "interventions": the drug(s)/device(s)/procedure(s) given in this arm, each with "name",
  "type" (e.g. "Drug", "Device", "Procedure", "Biological"), and dose/route/frequency if
  stated (leave any of those as empty string if not stated -- do not guess or infer a value
  that isn't actually written).

One array entry per genuinely distinct arm -- do not split a single arm into multiple entries
just because it's described in more than one place in the document; consolidate what you find
into one entry per arm. If no arms/groups section exists, return an empty list.

Respond with JSON only, matching the schema."""


OBJECTIVES_SCHEMA = {
    "type": "object",
    "properties": {
        "objectives": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "level": {"type": "string"},
                    "description": {"type": "string"},
                    "endpoints": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {"description": {"type": "string"}},
                            "required": ["description"],
                        },
                    },
                },
                "required": ["level", "description", "endpoints"],
            },
        },
    },
    "required": ["objectives"],
}

OBJECTIVES_PROMPT = """Find the Objectives and Endpoints section of this clinical trial
protocol (often a table titled "Objectives and Endpoints," "Study Objectives," or similar,
listing Primary/Secondary/Exploratory objectives each paired with their endpoints).

For each objective, extract:
- "level": "Primary", "Secondary", or "Exploratory" (use exactly one of these three, matching
  how the protocol itself categorizes it).
- "description": the objective's text exactly as written.
- "endpoints": the endpoint(s) listed under that specific objective, each as its own
  {"description": "..."} entry with the endpoint text exactly as written. Do not merge
  multiple endpoints into one entry.

Keep the protocol's own primary/secondary/exploratory grouping -- do not re-categorize an
objective based on your own judgment of its importance. If no such section exists, return an
empty list.

Respond with JSON only, matching the schema."""


def _call_gemini_pdf_json(pdf_bytes, prompt_text, schema):
    response = client.models.generate_content(
        model="gemini-2.5-pro",
        contents=[{"parts": [
            {"inline_data": {"mime_type": "application/pdf", "data": pdf_bytes}},
            {"text": prompt_text},
        ]}],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_json_schema=schema,
            thinking_config=types.ThinkingConfig(thinking_budget=-1),
        ),
    )
    text = (response.text or "").strip()
    return json.loads(text) if text else None


def extract_eligibility(pdf_bytes):
    data = _call_gemini_pdf_json(pdf_bytes, ELIGIBILITY_PROMPT, ELIGIBILITY_SCHEMA) or {}
    return {
        "inclusionCriteria": [str(c).strip() for c in data.get("inclusionCriteria", [])],
        "exclusionCriteria": [str(c).strip() for c in data.get("exclusionCriteria", [])],
    }


def extract_arms(pdf_bytes):
    data = _call_gemini_pdf_json(pdf_bytes, ARMS_PROMPT, ARMS_SCHEMA) or {}
    arms = []
    for a in data.get("arms", []):
        arms.append({
            "name": str(a.get("name", "")).strip(),
            "type": str(a.get("type", "")).strip(),
            "description": str(a.get("description", "")).strip(),
            "interventions": [
                {
                    "name": str(iv.get("name", "")).strip(),
                    "type": str(iv.get("type", "")).strip(),
                    "dose": str(iv.get("dose", "")).strip(),
                    "route": str(iv.get("route", "")).strip(),
                    "frequency": str(iv.get("frequency", "")).strip(),
                }
                for iv in a.get("interventions", [])
            ],
        })
    return arms


def extract_objectives(pdf_bytes):
    data = _call_gemini_pdf_json(pdf_bytes, OBJECTIVES_PROMPT, OBJECTIVES_SCHEMA) or {}
    objectives = []
    for o in data.get("objectives", []):
        objectives.append({
            "level": str(o.get("level", "")).strip(),
            "description": str(o.get("description", "")).strip(),
            "endpoints": [
                {"description": str(e.get("description", "")).strip()}
                for e in o.get("endpoints", [])
            ],
        })
    return objectives


def extract_protocol(pdf_path):
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    print("Extracting eligibility criteria...")
    eligibility = extract_eligibility(pdf_bytes)
    print(f"  {len(eligibility['inclusionCriteria'])} inclusion, "
          f"{len(eligibility['exclusionCriteria'])} exclusion criteria found")

    print("Extracting study arms/interventions...")
    arms = extract_arms(pdf_bytes)
    print(f"  {len(arms)} arm(s) found")

    print("Extracting objectives & endpoints...")
    objectives = extract_objectives(pdf_bytes)
    total_endpoints = sum(len(o["endpoints"]) for o in objectives)
    print(f"  {len(objectives)} objective(s), {total_endpoints} endpoint(s) found")

    return {
        "eligibilityCriteria": eligibility,
        "studyArms": arms,
        "objectives": objectives,
    }


if __name__ == "__main__":
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else input("Enter path to protocol PDF: ").strip()
    data = extract_protocol(pdf_path)
    out_path = "protocol_data.json"
    with open(out_path, "w") as f:
        json.dump(data, f, indent=2)
    print(f"Saved to {out_path}")