import json
import logging
from unittest.mock import patch

import pytest

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from app.src.JBGQualityGuards import summarize_quality_guards
from app.src.JBGModelAssessment import summarize_assessments


ASSESSMENT = dict(priority="minor", category="clarity", confidence="high",
                  meaning_risk="low", requires_rewrite=False)


def element(text, element_id="p1", kind="paragraph", **metadata):
    return dict(type=kind, element_id=element_id, text=text, **metadata)


def suggestor(elements, assess_changes=True):
    ai = JBGLangImprovSuggestorAI("test", "test", "base", 0,
                                 logging.getLogger(__name__), assess_changes=assess_changes)
    ai.json_structured_document = {"type": "docx", "elements": elements}
    return ai


def proposal(old, new, element_id="p1", kind="paragraph", assessment=ASSESSMENT):
    raw = dict(type=kind, element_id=element_id, old=old, new=new,
               motivation="Språklig förbättring.")
    if assessment is not None:
        raw["model_assessment"] = assessment
    return raw


def review(text, old, new, *, assessment=ASSESSMENT, assess_changes=True):
    ai = suggestor([element(text)], assess_changes)
    s = ai._validate_docx_suggestion(proposal(old, new, assessment=assessment))
    return ai._filter_with_quality_guards([s])[1][0]


@pytest.mark.parametrize("text,old,new,code", [
    ("De ska representera alla och fungera som beställare.", "ska representera", "representerar", "modality"),
    ("Det vore svårt.", "vore", "är", "modality"),
    ("Det borde följas upp.", "borde", "bör", "modality"),
    ("Upp till omkring två tredjedelar återstår.", "Upp till omkring", "Omkring", "quantifier"),
    ("Högst 10 ärenden.", "Högst", "Minst", "quantifier"),
    ("Det är inte tillräckligt.", "inte tillräckligt", "tillräckligt", "negation"),
    ("Villkoret gäller A eller B.", "eller", "och", "disjunction"),
    ("Mer strukturerad information ger bättre stöd.", "Mer strukturerad", "Strukturerad", "comparative"),
    ("Fördelningen av riskarbetet varierar.", "Fördelningen av riskarbetet", "Riskarbetet", "analytic_scope"),
    ("Omfattningen av felen som hittas varierar.", "felen som hittas", "felen", "scope_clause"),
    ("Omfattningen av felen som har upptäckts varierar.", "felen som har upptäckts", "felen", "scope_clause"),
    ("Andelen är 57 procent.", "57", "75", "numeric_fact"),
    ("Andelen är 57 procent.", "57 procent", "57", "numeric_fact"),
    ("Värdet är -7.", "-7", "7", "numeric_fact"),
    ("Intervallet är 6–700.", "6–700", "6 700", "numeric_fact"),
    ("Beslutet fattades 2024-09-24.", "2024-09-24", "25 september 2024", "numeric_fact"),
    ("A-kassornas inventeringar visar variation.", "A-kassornas", "Arbetslöshetskassornas", "terminology_direction"),
    ("Den nya arbetslöshetsförsäkringen gäller.", "arbetslöshetsförsäkringen", "försäkringen", "insurance_term"),
    ("Arbetslöshetskassorna (a-kassorna) har svarat.", "Arbetslöshetskassorna (a-kassorna)", "A-kassorna", "term_definition"),
    ("Systemen är regelstyrda automationer.", "regelstyrda automationer", "regelstyrd automation", "automation_number"),
    ("Systemet tas i produktion.", "produktion", "drift", "deployment_term"),
    ("Verktyget körs på deras egen server.", "deras", "en", "ownership"),
    ("Handläggarna följer enligt Arbetsförmedlingen ofta rekommendationerna.", "Arbetsförmedlingen", "IAF", "attribution"),
    ("Beslutet avgörs av handläggaren. Handläggarna följer enligt Arbetsförmedlingen ofta rekommendationerna.",
     "avgörs av handläggaren. Handläggarna följer enligt Arbetsförmedlingen", "avgörs enligt Arbetsförmedlingen av handläggaren. Handläggarna följer", "attribution"),
    ("Systemet heter ”Våra sidor”.", "Våra sidor", "Våra webbsidor", "quoted_text"),
    ("Källa: systeminventeringarna.", "systeminventeringarna", "inventeringarna", "source_label"),
    ("TODO: Här ska vi lägga till figur 6-7.", "6-7", "6–7", "internal_note"),
    ("Här vill vi lägga till ett avsnitt.", "Här vill vi", "TODO:", "internal_note"),
    ("Analyserna görs internt. Det beror bland annat på att lokala moduler påverkar riskbilden.",
     "internt", "internt. Det beror bland annat på att lokala moduler påverkar riskbilden", "context_duplication"),
    ("Analyserna görs internt. Det beror bland annat på lokala moduler.", "Det beror", "Analyserna görs internt. Det beror", "context_duplication"),
])
def test_blocks_known_risks_despite_high_confidence(text, old, new, code):
    row = review(text, old, new)
    assert not row.accepted
    assert "guard_" + code in {i.code for i in row.issues}


@pytest.mark.parametrize("text,old,new", [
    ("De har behov av stöd.", "har behov av", "behöver"),
    ("Det går att hantera frågan.", "går att hantera", "kan hanteras"),
    ("Systemen tilldelas riskpoäng.", "tilldelas", "får"),
    ("De skall följa upp.", "skall", "ska"),
    ("Samtliga system omfattas.", "Samtliga", "Alla"),
    ("Omkring 10 procent.", "Omkring", "Cirka"),
    ("Endast 10 procent.", "Endast", "Bara"),
    ("De faktorerna med högst påverkan.", "faktorerna med högst påverkan", "faktorer som påverkar mest"),
    ("De beskrivs i mer detalj nedan.", "i mer detalj", "mer i detalj"),
    ("Granskningen begränsas till AI.", "begränsas till", "omfattar endast"),
    ("Systemen som har identifierats redovisas.", "som har identifierats", "som identifierats"),
    ("Felen undgår upptäckt.", "undgår upptäckt", "upptäcks inte"),
    ("Värdet är 10 procent.", "10 procent", "10 %"),
    ("Intervallet är 6-700.", "6-700", "6–700"),
    ("Värdet är 1 000.", "1 000", "1000"),
    ("Beslutet fattades 2024-09-24.", "2024-09-24", "den 24 september 2024"),
    ("Figur 7 visar andelen. Figur 7 visar också risknivån.", "Figur 7 visar också", "Figuren visar också"),
    ("Resultatet beräknas enligt i förväg bestämda tröskelvärden.", "enligt i förväg bestämda tröskelvärden", "enligt tröskelvärden som bestämts i förväg"),
    ("Resultatet fördelar sig enligt följande:", "enligt följande", "så här"),
    ("Handläggarna följer enligt Arbetsförmedlingen ofta rekommendationerna.",
     "Handläggarna följer enligt Arbetsförmedlingen", "Enligt Arbetsförmedlingen följer handläggarna"),
    ("Arbetslöshetskassorna har svarat.", "Arbetslöshetskassorna", "A-kassorna"),
    ("Automationerna testas.", "Automationerna", "Automationer"),
    ("Vi har redan en upprepning. Vi har redan en upprepning. Det skedde sedan.", "skedde", "hände"),
    ("Systemet ”Våra sidor” utvecklas av leverantören.", "utvecklas av", "byggs av"),
])
def test_does_not_block_known_safe_edits(text, old, new):
    row = review(text, old, new)
    assert row.accepted, [i.code for i in row.issues]


@pytest.mark.parametrize("risk", ["medium", "high"])
@pytest.mark.parametrize("priority", ["critical", "minor"])
def test_model_risk_is_independent_of_priority(risk, priority):
    row = review("Det skedde i går.", "skedde", "hände",
                 assessment=dict(ASSESSMENT, meaning_risk=risk, priority=priority))
    assert not row.accepted
    assert "guard_model_meaning_risk" in {i.code for i in row.issues}


@pytest.mark.parametrize("assessment", [None, {}, dict(ASSESSMENT, meaning_risk="unknown")])
def test_missing_or_invalid_assessment_keeps_deterministic_fences(assessment):
    row = review("Det ska följas upp.", "ska", "bör", assessment=assessment)
    assert not row.accepted
    assert "guard_modality" in {i.code for i in row.issues}
    assert "guard_assessment_unavailable" in {i.code for i in row.issues}
    safe = review("Det skedde i går.", "skedde", "hände", assessment=assessment)
    assert safe.accepted
    assert "guard_assessment_unavailable" in {i.code for i in safe.issues}


def test_legacy_still_has_deterministic_guards_without_model_risk_gate():
    assert not review("Det ska ske.", "ska", "bör", assess_changes=False).accepted
    assert review("Det skedde.", "skedde", "hände", assess_changes=False,
                  assessment=dict(ASSESSMENT, meaning_risk="high")).accepted


def test_system_name_column_uses_source_metadata_not_model_category():
    els = [element("System", "h1", "table_cell", table_index=4, row_index=1, col_index=1),
           element("Funktion", "h2", "table_cell", table_index=4, row_index=1, col_index=2),
           element("MIR - utskick fullmakt önskad inträdesdatum", "s1", "table_cell", table_index=4, row_index=2, col_index=1),
           element("Utskick fullmakt", "d1", "table_cell", table_index=4, row_index=2, col_index=2)]
    ai = suggestor(els)
    raw = [proposal("utskick fullmakt", "utskick av fullmakt", "s1", "table_cell"),
           proposal("Utskick fullmakt", "Utskick av fullmakt", "d1", "table_cell")]
    accepted, reviewed = ai._filter_with_quality_guards([ai._validate_docx_suggestion(x) for x in raw])
    assert [s.element_id for s in accepted] == ["d1"]
    assert "guard_system_name" in {i.code for i in reviewed[0].issues}


def test_guard_does_not_guess_system_column_without_header():
    ai = suggestor([element("Utskick fullmakt", "s1", "table_cell", table_index=4, row_index=2, col_index=1)])
    s = ai._validate_docx_suggestion(proposal("Utskick fullmakt", "Utskick av fullmakt", "s1", "table_cell"))
    assert ai._filter_with_quality_guards([s])[0] == [s]


def test_full_pipeline_rejects_risk_and_persists_reason(tmp_path):
    ai = suggestor([element("De ska representera alla. Det skedde i går.")])
    raw = [proposal("ska representera", "representerar"), proposal("skedde", "hände")]
    # Run the actual single-call and batch paths without any external API.
    response = json.dumps(raw)
    with patch("app.src.JBGLangImprovSuggestorAI.create_openai_client", return_value=object()), \
         patch.object(ai, "_call_model", return_value=response), \
         patch.object(ai, "_call_model_for_chunks", return_value=[response]):
        for run in (ai.suggest_changes, ai.suggest_changes_token_aware_batching):
            run()
            assert [s.old for s in ai.validated_suggestions] == ["skedde"]
            path = tmp_path / "filter.json"
            ai.save_filter_report(path)
            rows = json.loads(path.read_text(encoding="utf-8"))
            assert any(not row["accepted"] and any(i["code"] == "guard_modality" for i in row["issues"]) for row in rows)
            stats = summarize_quality_guards(ai.filtered_review, True)
            assert stats["reviewed"] == 2 and stats["rejected"] == 1
            assert stats["by_reason"] == {"guard_modality": 1}
            assert summarize_assessments(raw, ai.validated_suggestions, risk_filter_active=True)["mode"] == "priority_shadow_risk_active"


def test_repeated_anchor_cannot_bypass_protected_context():
    row = review("En jämförelse med mer strukturerad information. Vi behöver mer hjälp.", "mer", "")
    assert not row.accepted
    assert "guard_comparative" in {i.code for i in row.issues}


def test_unique_or_changed_figure_reference_is_not_ignored():
    assert not review("Figur 7 visar andelen.", "Figur 7", "Figuren").accepted
    assert not review("Figur 7 visar andelen. Figur 7 visar risken.", "Figur 7 visar risken", "Figur 8 visar risken").accepted
