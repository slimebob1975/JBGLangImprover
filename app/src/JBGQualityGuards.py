"""Conservative, level-independent safeguards for local language proposals.

These are fences, not a semantic-equivalence classifier. Inspect the complete
source element and the result of the actual anchored replacement, rather than
trusting the model's motivation or examining a minimized span in isolation.
"""
from collections import Counter
from dataclasses import dataclass
import re


@dataclass(frozen=True)
class GuardIssue:
    code: str
    message: str
    severity: str = "error"


def _pattern(expr):
    return re.compile(expr, re.IGNORECASE)


# Explicit lexical markers only. A change to one of these is not proof of a
# semantic error; it is a reason to leave the decision to the human reviewer.
_MODALITY = _pattern(r"\b(?:ska|skall|skulle|bör|borde|måste|kan|kunde|behöver|behövde|vore|"
                     r"har\s+behov\s+av|hade\s+behov\s+av|går\s+att|gick\s+att)\b")
_MODAL_EQUIVALENTS = {"skall": "ska", "har behov av": "behöver", "hade behov av": "behövde",
                      "går att": "kan", "gick att": "kunde"}
_QUANTITY = _pattern(
    r"\b(?:upp\s+till|åtminstone|minst|högst(?=\s+(?:\d|en\b|ett\b|två\b|tre\b|fyra\b|fem\b|sex\b|sju\b|åtta\b|nio\b|tio\b))|som\s+mest|som\s+minst|"
    r"omkring|cirka|ungefär|drygt|knappt|nästan|samtliga|alla|varje|"
    r"endast|bara|enbart|begränsas\s+till|är\s+begränsad\s+till|vissa|sannolikt|eventuellt|möjligen|tillräckligt)\b"
)
_QUANTITY_EQUIVALENTS = {
    "samtliga": "alla", "endast": "bara", "enbart": "bara",
    "cirka": "omkring", "ungefär": "omkring", "åtminstone": "minst",
    "som mest": "högst", "som minst": "minst", "begränsas till": "bara", "är begränsad till": "bara",
}
_NEGATION = _pattern(r"\b(?:inte|ej|icke|aldrig|ingen|inget|inga|utan)\b")
_NEGATION_EQUIVALENTS = {"ej": "inte", "icke": "inte", "ingen": "inte", "inget": "inte", "inga": "inte"}
# Well-defined equivalence for a common plain-language rewrite. Do not turn
# this into a general permission to add or remove negation.
_NEGATION_REWRITE = _pattern(r"\b(?:undgår\s+upptäckt|upptäcks\s+inte)\b")
_DISJUNCTION = _pattern(r"\b(?:eller|alternativt|antingen)\b")
_COMPARATIVE = _pattern(r"\b(?:mer|mindre|fler|färre)\s+([a-zåäö][a-zåäö-]*)\b")
_SCOPE = _pattern(
    r"\b(?:som\s+(?:hittas|upptäcks|redovisas|identifieras|observeras|registreras|"
    r"(?:har\s+)?(?:hittats|upptäckts|redovisats|identifierats|observerats|registrerats)))\b"
)
_ANALYTIC_SCOPE = _pattern(r"\b(?:fördelningen|omfattningen|andelen)\s+av\b")
_LONG_KASSA = _pattern(r"\barbetslöshets?kass(?:a|an|ans|as|or|orna|ors|ornas)\b")
_INSURANCE = _pattern(r"\barbetslöshetsförsäkring\w*\b")
_AUTOMATION_SINGULAR = _pattern(r"\bautomation(?:en|ens|s)?\b")
_AUTOMATION_PLURAL = _pattern(r"\bautomationer(?:na|nas|s)?\b")
_DEFINITION = _pattern(r"\barbetslöshets?kass(?:a|an|or|orna)\s*\(\s*a[-\u2011]kass(?:a|an|or|orna)\s*\)")
_DEPLOYMENT = _pattern(r"\bi\s+(?:produktion|drift|bruk)\b")
_OWNERSHIP = _pattern(r"\b(?:deras|dess|vår|vårt|våra|sin|sitt|sina|hans|hennes)\s+egen\b|"
                     r"\b(?:deras|dess|vår|vårt|våra|sin|sitt|sina|hans|hennes)\s+eget\b|"
                     r"\b(?:deras|dess|vår|vårt|våra|sin|sitt|sina|hans|hennes)\s+egna\b")
_ATTRIBUTION = _pattern(r"\benligt\b(?!\s+följande\b)")
_ATTRIBUTION_SOURCE = _pattern(r"\benligt\s+(?!följande\b)([a-zåäö][\w-]*)")
_WORD = re.compile(r"[^\W_]+(?:[-\u2011][^\W_]+)*", re.UNICODE)
_SENTENCE_SPLIT = re.compile(r"(?<=[.!?])\s+|\n+")
_NOTE = _pattern(r"\b(?:TODO|FIXME|XXX|TK)\b")
_SOURCE_LABEL = _pattern(r"^\s*(?:källa|källor)\s*:")
_QUOTED = re.compile(r'”([^”\n]+)”|“([^”\n]+)”|"([^"\n]+)"|»([^«\n]+)«')
_SYSTEM_HEADERS = {"system", "systemnamn", "systemets namn", "namn på system",
                   "namn på tekniskt system", "systembenämning"}

# Numbers are compared in order, including signs, decimals and percent signs.
# Common Swedish digit grouping, range dashes and a space before % are handled
# as typography. ISO/written date equivalence is delegated to the existing
# date parser by the caller, so it cannot be inferred from model motivation.
_NUMBER = re.compile(r"<RANGE>|[+\-\u2212]?[0-9]+(?:[ \u00a0\u202f][0-9]{3}(?![0-9]))*(?:[,.][0-9]+)?(?:[ \u00a0\u202f]*%)?")


def _normalized(value):
    return " ".join(value.casefold().replace("\u2011", "-").split())


def _matches(pattern, text, equivalents=None):
    values = [_normalized(match.group(0)) for match in pattern.finditer(text)]
    return Counter((equivalents or {}).get(value, value) for value in values)


def _words(text):
    return [_normalized(m.group(0)) for m in _WORD.finditer(text)]


def _number_signature(text):
    # Preserve the distinction between a range, two numbers, a sign and a
    # thousands separator. Only the glyph/spacing of a range dash is ignored.
    text = re.sub(r"(?<=\d)\s*[-\u2010\u2011\u2012\u2013\u2014]\s*(?=[+\-\u2212]?\d)", " <RANGE> ", text)
    text = re.sub(r"(?<=\d)\s+procent\b", "%", text, flags=re.I)
    return tuple(re.sub(r"\s+", "", m.group(0)).replace("\u2212", "-")
                 for m in _NUMBER.finditer(text))


def _boundary_overlap(left, right):
    """Length in whole words of an adjacent repeated phrase (both sides)."""
    a, b = _words(left), _words(right)
    for size in range(min(len(a), len(b)), 2, -1):
        if a[-size:] == b[:size] and sum(map(len, b[:size])) >= 12:
            return size
    return 0


def _repeated_sentences(text):
    sentences = [tuple(_words(part)) for part in _SENTENCE_SPLIT.split(text)]
    return Counter(a for a, b in zip(sentences, sentences[1:]) if a == b and len(a) >= 4)


def _introduces_duplication(source, start, end, replacement):
    left, old, right = source[:start], source[start:end], source[end:]
    if (_boundary_overlap(left, replacement + right) > _boundary_overlap(left, old + right)
            or _boundary_overlap(left + replacement, right) > _boundary_overlap(left + old, right)):
        return True
    return bool(_repeated_sentences(left + replacement + right) - _repeated_sentences(source))


def _attribution_moved(before, after):
    """Detect changes to explicit 'enligt' and likely cross-claim relocation.

    Compare the lexical surroundings in the sentence, not its absolute index:
    splitting an earlier sentence must not itself change an attribution.
    This is intentionally a heuristic, not a Swedish dependency parser.
    """
    def contexts(text):
        result = []
        for sentence in _SENTENCE_SPLIT.split(text):
            if _ATTRIBUTION.search(sentence):
                words = set(_words(sentence))
                result.extend([words] * len(list(_ATTRIBUTION.finditer(sentence))))
        return result

    # A changed source word is itself grounds for caution, even if the
    # rest of the sentence is unchanged (e.g. another named authority).
    if [w.casefold() for w in _ATTRIBUTION_SOURCE.findall(before) if w[0].isupper()] != [w.casefold() for w in _ATTRIBUTION_SOURCE.findall(after) if w[0].isupper()]:
        return True
    old, new = contexts(before), contexts(after)
    if len(old) != len(new):
        return True
    # For each occurrence, preserve the surrounding claim sufficiently.
    # A simple reordering within that same claim is allowed.
    for a, b in zip(old, new):
        if len(a & b) / max(1, len(a | b)) < 0.5:
            return True
    return False


def _quote_contents(text):
    return Counter(next(group for group in m.groups() if group is not None)
                   for m in _QUOTED.finditer(text))


_FIGURE_REFERENCE = _pattern(r"\b(figur|tabell)\s+(\d+)\b")


def _number_shift(before, after, normalize_dates):
    # A repeated explicit reference may become 'figuren'/'tabellen' when
    # the same unique reference is still explicit earlier in this element.
    # Never discard a unique reference, or a reference in a mixed-ID context.
    old_refs = list(_FIGURE_REFERENCE.finditer(before))
    new_refs = list(_FIGURE_REFERENCE.finditer(after))
    if old_refs and len(old_refs) > len(new_refs) >= 1:
        identities = {(_normalized(m.group(1)), m.group(2)) for m in old_refs + new_refs}
        if len(identities) == 1:
            kind, _ = next(iter(identities))
            definite = "figuren" if kind == "figur" else "tabellen"
            count = len(old_refs) - len(new_refs)
            if len(re.findall(r"\b" + definite + r"\b", after, re.I)) - len(re.findall(r"\b" + definite + r"\b", before, re.I)) == count:
                # Keep the initial explicit occurrence. This only exempts
                # repeated references, never numeric values elsewhere.
                for match in reversed(old_refs[-count:]):
                    before = before[:match.start()] + definite + before[match.end():]
    return _number_signature(normalize_dates(before)) != _number_signature(normalize_dates(after))


def _scope_signature(text):
    return Counter(re.sub(r"^som har ", "som ", _normalized(m.group(0)))
                   for m in _SCOPE.finditer(text))


class QualityGuards:
    """Check proposals against the source structure; never rewrite them."""

    def __init__(self, elements):
        self.elements = {e.get("element_id"): e for e in elements}
        self.system_columns = set()
        # Use explicit column headers, never specific report/table IDs or
        # the model's classification. Only actual multi-column tables qualify.
        first_row_columns = {}
        for e in self.elements.values():
            if e.get("type") == "table_cell" and e.get("row_index") == 1:
                first_row_columns.setdefault(e.get("table_index"), set()).add(e.get("col_index"))
        for e in self.elements.values():
            table, col = e.get("table_index"), e.get("col_index")
            if (e.get("type") == "table_cell" and e.get("row_index") == 1
                    and isinstance(table, int) and isinstance(col, int)
                    and len(first_row_columns.get(table, ())) > 1
                    and _normalized(e.get("text") or "").rstrip(":") in _SYSTEM_HEADERS):
                self.system_columns.add((table, col))

    def review(self, suggestion, find_span, normalize_dates, assess_changes=False):
        issues = {}

        def add(code, message, severity="error"):
            issues[code] = GuardIssue(code, message, severity)

        if assess_changes:
            assessment = suggestion.model_assessment
            if suggestion.assessment_status == "valid" and assessment is not None:
                if assessment.meaning_risk in {"medium", "high"}:
                    add("guard_model_meaning_risk", "Modellen anger förhöjd betydelserisk; kräver manuell bedömning.")
            else:
                add("guard_assessment_unavailable", "Giltig modellbedömning saknas; endast oberoende skydd kan tillämpas.", "warning")

        element = self.elements.get(suggestion.element_id)
        if element is None:
            add("guard_source_unavailable", "Källelement saknas vid kvalitetskontroll.")
            return list(issues.values())
        source = element.get("text") or ""
        if _SOURCE_LABEL.search(source):
            add("guard_source_label", "Skyddar figurens eller tabellens källangivelse.")
        if (_NOTE.search(source) or _NOTE.search(suggestion.new)):
            add("guard_internal_note", "Skyddar interna arbetsmarkörer mot lokala textändringar.")
        if (element.get("type") == "table_cell" and (element.get("row_index") or 0) > 1
                and (element.get("table_index"), element.get("col_index")) in self.system_columns):
            add("guard_system_name", "Skyddar innehållet i en uttryckligen namngiven systemkolumn.")

        # The existing matcher may choose a particular repeated anchor. Test
        # every exact occurrence, so a protected occurrence cannot be bypassed
        # by relying on a different occurrence's benign context.
        spans = [(m.start(), m.start() + len(suggestion.old)) for m in re.finditer(
            "(?=" + re.escape(suggestion.old) + ")", source)] if suggestion.old else []
        if not spans:
            span = find_span(suggestion.old, source)
            if span is not None:
                spans = [span]
        if not spans:
            add("guard_anchor_unavailable", "Ändringen kan inte förankras vid kvalitetskontroll.")
            return list(issues.values())

        for start, end in spans:
            before = source
            after = source[:start] + suggestion.new + source[end:]
            if before == after:
                continue
            if _matches(_MODALITY, before, _MODAL_EQUIVALENTS) != _matches(_MODALITY, after, _MODAL_EQUIVALENTS):
                add("guard_modality", "Ändrar en uttrycklig modalitetsmarkör.")
            if _matches(_QUANTITY, before, _QUANTITY_EQUIVALENTS) != _matches(_QUANTITY, after, _QUANTITY_EQUIVALENTS):
                add("guard_quantifier", "Ändrar en uttrycklig kvantifiering eller osäkerhetsmarkör.")
            # Check negation after a narrow, known equivalent rewrite.
            neg_before = _NEGATION_REWRITE.sub("undgår upptäckt", before)
            neg_after = _NEGATION_REWRITE.sub("undgår upptäckt", after)
            if _matches(_NEGATION, neg_before, _NEGATION_EQUIVALENTS) != _matches(_NEGATION, neg_after, _NEGATION_EQUIVALENTS):
                add("guard_negation", "Ändrar en uttrycklig negation.")
            if _matches(_DISJUNCTION, before) != _matches(_DISJUNCTION, after):
                add("guard_disjunction", "Ändrar ett uttryckligt alternativ eller en disjunktion.")
            # Do not treat all stylistic uses of mer/mindre as protected:
            # flag loss of a degree marker when its following word survives.
            comparison_before = re.sub(r"\bi\s+mer\s+detalj\b", "mer i detalj", before, flags=re.I)
            comparison_after = re.sub(r"\bi\s+mer\s+detalj\b", "mer i detalj", after, flags=re.I)
            for m in _COMPARATIVE.finditer(comparison_before):
                phrase = _normalized(m.group(0))
                if (_matches(_COMPARATIVE, comparison_before)[phrase] > _matches(_COMPARATIVE, comparison_after)[phrase]
                        and m.group(1).casefold() not in {"i", "än", "eller"}
                        and re.search(r"\b" + re.escape(m.group(1)) + r"\b", after, re.I)):
                    add("guard_comparative", "En jämförelsemarkör försvinner framför ett kvarvarande ord.")
            if _matches(_ANALYTIC_SCOPE, before) - _matches(_ANALYTIC_SCOPE, after):
                add("guard_analytic_scope", "En uttrycklig analytisk avgränsning försvinner.")
            if _scope_signature(before) - _scope_signature(after):
                add("guard_scope_clause", "En uttrycklig avgränsning av identifierade eller redovisade objekt försvinner.")
            if _number_shift(before, after, normalize_dates):
                add("guard_numeric_fact", "Ändrar ordningsföljd, värde eller tecken för en sifferuppgift.")
            if _matches(_LONG_KASSA, after) - _matches(_LONG_KASSA, before):
                add("guard_terminology_direction", "Inför en lång a-kassa-form i strid med den fasta terminologiregeln.")
            if _matches(_INSURANCE, before) != _matches(_INSURANCE, after):
                add("guard_insurance_term", "Ändrar den skyddade termen arbetslöshetsförsäkring.")
            if _matches(_DEFINITION, before) - _matches(_DEFINITION, after):
                add("guard_term_definition", "Tar bort en uttrycklig introduktion av a-kassa-beteckningen.")
            singular_delta = len(_AUTOMATION_SINGULAR.findall(after)) - len(_AUTOMATION_SINGULAR.findall(before))
            plural_delta = len(_AUTOMATION_PLURAL.findall(after)) - len(_AUTOMATION_PLURAL.findall(before))
            if singular_delta * plural_delta < 0:
                add("guard_automation_number", "Byter mellan automation som företeelse och automationer som objekt.")
            if _matches(_DEPLOYMENT, before) - _matches(_DEPLOYMENT, after) and _matches(_DEPLOYMENT, after) - _matches(_DEPLOYMENT, before):
                add("guard_deployment_term", "Byter teknisk beteckning för införande eller produktionssättning.")
            if _matches(_OWNERSHIP, before) - _matches(_OWNERSHIP, after):
                add("guard_ownership", "En uttrycklig ägande- eller ansvarskoppling vid egen/egna förändras.")
            if _attribution_moved(before, after):
                add("guard_attribution", "En uttrycklig enligt-hänvisning ändras eller kan ha flyttats till ett annat påstående.")
            if _quote_contents(before) - _quote_contents(after):
                add("guard_quoted_text", "Ändrar befintlig text inom citattecken.")
            if _introduces_duplication(source, start, end, suggestion.new):
                add("guard_context_duplication", "Ändringen skapar en upprepning av angränsande text.")

        return list(issues.values())


def summarize_quality_guards(reviewed, assess_changes):
    rows = [row for row in reviewed if any(i.code.startswith("guard_") for i in row.issues)]
    reasons = Counter(i.code for row in rows for i in row.issues
                      if i.code.startswith("guard_") and i.severity == "error")
    return {"mode": "active", "model_risk_filter": "active" if assess_changes else "not_available",
            "reviewed": len(reviewed),
            "rejected": sum(any(i.code.startswith("guard_") and i.severity == "error"
                                for i in row.issues) for row in rows),
            "by_reason": dict(sorted(reasons.items())),
            "assessment_unavailable": sum(any(i.code == "guard_assessment_unavailable"
                                               for i in row.issues) for row in rows)}
