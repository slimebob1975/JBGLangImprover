"""Model assessments; low risk never overrides backend safety checks."""
from dataclasses import dataclass
from typing import Optional

try:
    from app.src.JBGEditLevel import PRIORITY_DEFINITIONS
except ImportError:
    from JBGEditLevel import PRIORITY_DEFINITIONS

PRIORITIES = tuple(PRIORITY_DEFINITIONS)
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
        "\n\n## Modellbedömning (oberoende av bearbetningsnivå)\n"
        "Behåll den befintliga JSON-listan och alla tidigare obligatoriska fält. "
        "Lägg till objektet model_assessment i varje redan motiverat ändringsförslag. "
        "Bedöm förslaget oberoende av vald bearbetningsnivå. "
        "Samma konkreta ingrepp ska få samma prioritet oberoende av nivå, kategori eller spannlängd. "
        "En obligatorisk terminologiregel avgör inte i sig prioriteten. "
        "Bedöm varje självständigt ingrepp för sig, inte nyttan av en samling ändringar. "
        "priority: " + "; ".join(f"{key} ({definition})" for key, definition in PRIORITY_DEFINITIONS.items()) + ". "
        "category: " + ", ".join(CATEGORIES) + ". "
        "confidence: high (välgrundad nytta och korrekthet), medium (viss osäkerhet), "
        "low (otillräckligt stöd). Hög säkerhet kräver kontroll av hela den ändrade formuleringen i sitt sammanhang. "
        "meaning_risk: low (ingen identifierad betydelseförskjutning efter kontroll), "
        "medium (rimlig möjlighet till betydelseförskjutning eller osäker tolkning), "
        "high (sannolik eller tydlig betydelseförskjutning). Bedöm särskilt modalitet, "
        "statistiska gränser, osäkerhetsmarkörer, villkor, fackbegrepp och rättsliga avgränsningar. "
        "Bedöm risken separat från confidence och priority; klassificera inte osäker tolkning som low. "
        "requires_rewrite: true om en större omformulering av meningsbyggnaden krävs, annars false. "
        'Exempel: "model_assessment": {"priority":"significant","category":"complex_expression",'
        '"confidence":"high","meaning_risk":"low","requires_rewrite":false}. '
        "Bedömningen är inte ett tillstånd att kringgå skyddsregler. "
        "Föreslå inte fler ändringar bara för att fylla kategorierna; likvärdiga stilvarianter ska fortfarande undvikas."
    )


def summarize_assessments(raw, accepted, *, risk_filter_active=False) -> dict:
    def count(items):
        counts = {"valid": 0, "missing": 0, "invalid": 0,
                  "by_priority": {p: 0 for p in PRIORITIES}}
        for assessment, status in items:
            counts[status] += 1
            if assessment is not None:
                counts["by_priority"][assessment.priority] += 1
        return counts
    return {"mode": "priority_shadow_risk_active" if risk_filter_active else "shadow", "raw": count(parse_assessment(item) for item in raw),
            "accepted": count((s.model_assessment, s.assessment_status) for s in accepted)}
