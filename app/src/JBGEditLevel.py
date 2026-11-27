"""Central definition of local review levels."""
from dataclasses import asdict, dataclass
from typing import Optional

MIN_LEVEL = 1
MAX_LEVEL = 100
DEFAULT_LEVEL = 70


@dataclass(frozen=True)
class EditProfile:
    max_level: int
    key: str
    name: str
    description: str


# Each band adds a criterion; the same definitions are used in assessment prompts.
PRIORITY_DEFINITIONS = {
    "critical": "Entydigt språkfel eller allvarligt hinder för korrekt förståelse.",
    "significant": "Betydande förbättring av begripligheten.",
    "moderate": "Tydlig och konkret klarspråksnytta.",
    "minor": "Liten men motiverad förbättring med konkret nytta.",
    "cosmetic": "Likvärdig stilvariant utan konkret förbättringsnytta.",
}
PROFILES = (
    EditProfile(20, "very_conservative", "Mycket varsam", "Föreslår rättelser av tydliga språkfel och formuleringar som är mycket svåra att förstå."),
    EditProfile(40, "conservative", "Varsam", "Föreslår också ändringar som gör texten betydligt lättare att förstå."),
    EditProfile(60, "balanced", "Balanserad", "Föreslår också enklare ord och tydligare meningar."),
    EditProfile(80, "thorough", "Noggrann", "Föreslår också mindre ändringar för mer exakta formuleringar, bättre flyt och färre onödiga ord."),
    EditProfile(100, "very_thorough", "Mycket noggrann", "Föreslår också små förbättringar av exempelvis skiljetecken och ordval, när de gör texten bättre."),
)



def get_edit_profile(level: Optional[int]) -> Optional[EditProfile]:
    if level is None:
        return None  # Legacy clients do not select a profile.
    if type(level) is not int or not MIN_LEVEL <= level <= MAX_LEVEL:
        raise ValueError("edit_level must be an integer from 1 to 100")
    return next(profile for profile in PROFILES if level <= profile.max_level)


def edit_level_config() -> dict:
    return {"min": MIN_LEVEL, "max": MAX_LEVEL, "default": DEFAULT_LEVEL,
            "profiles": [asdict(profile) for profile in PROFILES]}


# The two highest bands share the minor priority class; the last band adds
# scrutiny of subtle details, not permission for cosmetic changes.
_LEVEL_CRITERIA = (
    ("critical", PRIORITY_DEFINITIONS["critical"]),
    ("significant", PRIORITY_DEFINITIONS["significant"]),
    ("moderate", PRIORITY_DEFINITIONS["moderate"] + " Exempelvis ordval, syftning och meningsbyggnad."),
    ("minor", PRIORITY_DEFINITIONS["minor"] + " Exempelvis mer exakta formuleringar, bättre flyt och färre onödiga ord."),
    ("minor", "Även små förbättringar av skiljetecken, ordval och språkliga ojämnheter, endast när de gör texten bättre."),
)
PROMPT_PROFILES = {
    profile.key: "\n".join(f"- {criterion}" for _, criterion in _LEVEL_CRITERIA[:i + 1])
    for i, profile in enumerate(PROFILES)
}

# Shared safeguards, independent of level and advisory assessment.
REVIEW_GUARDRAILS = (
    "Bevara sakuppgifter, fackbegrepp och rättslig och statistisk innebörd, inklusive "
    "modalitet, negationer, villkor, avgränsningar, kvantifieringar och osäkerhetsmarkörer. "
    "Exempelvis har 'upp till' och 'omkring' olika funktioner; bevara båda när de förekommer. "
    "Gissa inte avsedd innebörd i oklar eller skadad text. "
    "En mer formell formulering är inte i sig en klarspråksförbättring. "
    "Separera oberoende ingrepp i små, exakt förankrade, icke överlappande förslag "
    "med tillräcklig kontext för entydig matchning. Håll obligatoriska terminologibyten "
    "och entydiga rättelser åtskilda från frivilliga förbättringar när det är möjligt. "
    "Dela inte en sammanhängande grammatisk ändring om delarna inte fungerar var för sig. "
    "Kontrollera att den ändrade texten är korrekt i sitt sammanhang och att motiveringen "
    "beskriver det faktiska ingreppet."
)


def build_local_prompt(base_prompt: str, level: Optional[int]) -> str:
    profile = get_edit_profile(level)
    if profile is None:
        return base_prompt
    return (
        base_prompt + "\n\n## Bearbetningsnivå (lokal granskning)\n"
        + f"Önskad ingreppsnivå: {level}/100 (1 = minst, 100 = mest ingående).\n"
        + f"Profil: {profile.name}.\n"
        + "Använd det exakta värdet även inom profilen: ju högre värde, desto mindre behöver "
        "den konkreta nyttan av en frivillig ändring vara för att du ska föreslå den. "
        "Profilen anger den övergripande inriktningen. Nivån ändrar inte kraven på säkerhet "
        "eller på nödvändiga rättelser.\n"
        + "Föreslå endast ändringar som uppfyller något av följande kumulativa kriterier:\n"
        + PROMPT_PROFILES[profile.key]
        + "\nObligatoriska terminologiregler och entydiga språkfel gäller på alla nivåer, "
        "förutsatt att ändringen är säker och inte berör skyddad text. "
        "Denna instruktion styr endast tröskeln för vilka nyttiga ändringar som föreslås. "
        "Alla skyddsregler, fasta terminologiregler, krav på bibehållen betydelse, "
        "regler för ändringarnas omfattning och JSON-format gäller oförändrat och har företräde. "
        "Gör inga ändringar enbart för att uppnå ett visst antal förslag. "
        "Lämna fungerande formuleringar utan konkret förbättringsmöjlighet orörda.\n"
        + REVIEW_GUARDRAILS
    )
