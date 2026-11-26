"""Central definition of local review levels; no prompt effect in patch 2."""
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
