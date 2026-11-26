"""Advisory model assessments; never a backend safety verdict."""
from dataclasses import dataclass
from typing import Optional

PRIORITIES = ("critical", "significant", "moderate", "minor", "cosmetic")
CATEGORIES = ("correctness", "clarity", "complex_expression", "sentence_structure",
              "precision", "redundancy", "terminology", "typography", "other")
CONFIDENCE = ("high", "medium", "low")
MEANING_RISKS = ("low", "medium", "high")


@dataclass(frozen=True)
class ModelAssessment:
    priority: str
    category: str
    confidence: str
    meaning_risk: str
    requires_rewrite: bool


def parse_assessment(item: dict) -> tuple[Optional[ModelAssessment], str]:
    if "model_assessment" not in item:
        return None, "missing"
    value = item["model_assessment"]
    if not isinstance(value, dict):
        return None, "invalid"
    for key, allowed in (("priority", PRIORITIES), ("category", CATEGORIES),
                         ("confidence", CONFIDENCE), ("meaning_risk", MEANING_RISKS)):
        if value.get(key) not in allowed:
            return None, "invalid"
    if type(value.get("requires_rewrite")) is not bool:
        return None, "invalid"
    return ModelAssessment(**{key: value[key] for key in ModelAssessment.__dataclass_fields__}), "valid"


def assessment_prompt() -> str:
    return (
        "\n\n## Modellbedömning (endast diagnostik)\n"
        "Behåll den befintliga JSON-listan och alla tidigare obligatoriska fält. "
        "Lägg till objektet model_assessment i varje redan motiverat ändringsförslag. "
        "Bedöm förslaget oberoende av vald bearbetningsnivå. "
        "priority: critical (uppenbart allvarligt språkfel eller förståelsehinder), "
        "significant (tydlig betydande förbättring), moderate (konkret normal förbättring), "
        "minor (liten konkret förbättring), cosmetic (enbart ytlig detalj). "
        "category: " + ", ".join(CATEGORIES) + ". "
        "confidence: high, medium eller low (säkerhet om nyttan och korrektheten). "
        "meaning_risk: low, medium eller high (risk för betydelseförskjutning). "
        "requires_rewrite: true om en större omformulering av meningsbyggnaden krävs, annars false. "
        'Exempel: "model_assessment": {"priority":"significant","category":"complex_expression",'
        '"confidence":"high","meaning_risk":"low","requires_rewrite":false}. '
        "Bedömningen är inte ett tillstånd att kringgå skyddsregler. "
        "Föreslå inte fler ändringar bara för att fylla kategorierna; likvärdiga stilvarianter ska fortfarande undvikas."
    )


def summarize_assessments(raw, accepted) -> dict:
    def count(items):
        counts = {"valid": 0, "missing": 0, "invalid": 0,
                  "by_priority": {p: 0 for p in PRIORITIES}}
        for assessment, status in items:
            counts[status] += 1
            if assessment is not None:
                counts["by_priority"][assessment.priority] += 1
        return counts
    return {"mode": "shadow", "raw": count(parse_assessment(item) for item in raw),
            "accepted": count((s.model_assessment, s.assessment_status) for s in accepted)}
