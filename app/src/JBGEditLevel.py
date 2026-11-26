"""Central definition of local review levels."""
from dataclasses import asdict, dataclass
from typing import Optional

MIN_LEVEL = 1
MAX_LEVEL = 100
DEFAULT_LEVEL = 50


@dataclass(frozen=True)
class EditProfile:
    max_level: int
    key: str
    name: str
    description: str


PROFILES = (
    EditProfile(20, "very_conservative", "Mycket varsam", "Åtgärdar bara tydliga och viktiga klarspråksproblem."),
    EditProfile(40, "conservative", "Varsam", "Åtgärdar tydliga problem och låter fungerande formuleringar vara."),
    EditProfile(60, "balanced", "Balanserad", "Förbättrar tydlighet, meningsbyggnad och läsbarhet när det behövs."),
    EditProfile(80, "thorough", "Noggrann", "Uppmärksammar även mindre språkliga brister och onödig komplexitet."),
    EditProfile(100, "very_thorough", "Mycket noggrann", "Granskar även subtila språkliga detaljer när ändringen ger tydlig nytta."),
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


# These instructions only adjust the threshold for useful local changes.
PROMPT_PROFILES = {
    "very_conservative": "Föreslå endast säkra, tydligt nödvändiga förbättringar: uppenbara språkfel eller betydande hinder för förståelsen. Avstå från mindre förbättringar.",
    "conservative": "Föreslå säkra förbättringar av tydliga språkfel och klarspråksproblem. Avstå från marginella förbättringar av redan fungerande text.",
    "balanced": "Följ den ordinarie klarspråkspolicyns tröskel för ändringar: föreslå förbättringar med tydlig nytta för begriplighet, tydlighet eller korrekthet.",
    "thorough": "Uppmärksamma även mindre men konkreta förbättringar av tydlighet, precision, meningsbyggnad och onödig komplexitet. Föreslå dem endast när nyttan är tydlig.",
    "very_thorough": "Granska även subtila språkproblem, onödiga småord och mindre hinder för läsningen. Föreslå även små förbättringar när de ger en konkret nytta; avstå från likvärdiga stilvarianter.",
}


def build_local_prompt(base_prompt: str, level: Optional[int]) -> str:
    profile = get_edit_profile(level)
    if profile is None:
        return base_prompt
    return (
        base_prompt + "\n\n## Bearbetningsnivå (lokal granskning)\n"
        + f"Profil: {profile.name}.\n" + PROMPT_PROFILES[profile.key]
        + "\nDenna instruktion styr endast tröskeln för vilka nyttiga ändringar som föreslås. "
        "Alla skyddsregler, fasta terminologiregler, krav på bibehållen betydelse, "
        "regler för ändringarnas omfattning och JSON-format gäller oförändrat och har företräde. "
        "Gör inga ändringar enbart för att uppnå ett visst antal förslag. "
        "Lämna fungerande formuleringar utan konkret förbättringsmöjlighet orörda."
    )
