from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Skipped:
    """A file or company that produced no presentation, with a reason a user can act on."""

    name: str
    reason: str


@dataclass
class GenerationResult:
    written: list[Path] = field(default_factory=list)
    skipped: list[Skipped] = field(default_factory=list)
