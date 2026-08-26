import json
from decimal import Decimal
from pathlib import Path
from typing import Optional

from pydantic import BaseModel, ConfigDict, model_validator

from hopptx.paths import RUNS_FILE, TEMPLATE_DIR


class InputConfig(BaseModel):
    model_config = ConfigDict(frozen=True)

    type: str  # "file" or "directory"
    path: str


class RunConfig(BaseModel):
    model_config = ConfigDict(frozen=True)

    name: str
    start: str
    end: str
    input: InputConfig
    exclude_canceled: bool
    cost_per_training: Decimal
    evaluation_rate_percentage: Decimal
    top_topics_count: int
    top_topics_hard_cutoff: int = 10
    group_topics: bool
    clean_topic_names: bool
    topic_fuzzy_match_threshold: int
    strip_company_prefix: bool
    category_reference: str
    category_fuzzy_match_threshold: int
    max_programs_per_table_page: int
    single_page_soft_overflow: int
    template: str
    topic_max_chars: Optional[int]
    topic_truncation_suffix: Optional[str]
    filename_suffix: Optional[str] = None
    companies_inline: Optional[list[str]] = None

    @model_validator(mode="after")
    def _check_input_type(self) -> "RunConfig":
        template_path = TEMPLATE_DIR / self.template
        if not template_path.exists():
            raise ValueError(
                f"Run '{self.name}': template file not found: {template_path}"
            )
        if self.input.type not in ("file", "directory"):
            raise ValueError(
                f"Run '{self.name}': input.type must be 'file' or 'directory', got '{self.input.type}'"
            )
        if self.max_programs_per_table_page <= 0:
            raise ValueError(
                f"Run '{self.name}': max_programs_per_table_page must be > 0"
            )
        if self.single_page_soft_overflow < 0:
            raise ValueError(
                f"Run '{self.name}': single_page_soft_overflow must be >= 0"
            )
        if self.cost_per_training < 0:
            raise ValueError(
                f"Run '{self.name}': cost_per_training must be >= 0"
            )
        if (
            self.evaluation_rate_percentage < 0
            or self.evaluation_rate_percentage > 100
        ):
            raise ValueError(
                f"Run '{self.name}': evaluation_rate_percentage must be between 0 and 100"
            )
        if self.top_topics_count <= 0:
            raise ValueError(
                f"Run '{self.name}': top_topics_count must be > 0"
            )
        if self.top_topics_hard_cutoff <= 0:
            raise ValueError(
                f"Run '{self.name}': top_topics_hard_cutoff must be > 0"
            )
        if self.topic_truncation_suffix is not None:
            if self.topic_truncation_suffix == "":
                raise ValueError(
                    f"Run '{self.name}': topic_truncation_suffix must be null or a non-empty string"
                )
            if self.topic_max_chars is None or self.topic_max_chars <= 0:
                raise ValueError(
                    f"Run '{self.name}': topic_max_chars must be > 0 when topic_truncation_suffix is set"
                )
        if self.filename_suffix is not None and self.filename_suffix.strip() == "":
            raise ValueError(
                f"Run '{self.name}': filename_suffix must be null or a non-empty string"
            )
        if self.companies_inline is not None:
            if len(self.companies_inline) == 0:
                raise ValueError(
                    f"Run '{self.name}': companies_inline must be null or a non-empty list"
                )
            for company in self.companies_inline:
                if company.strip() == "":
                    raise ValueError(
                        f"Run '{self.name}': companies_inline cannot contain empty company names"
                    )
        return self

    @property
    def template_path(self) -> Path:
        return TEMPLATE_DIR / self.template


def load_runs(runs_file: Path = RUNS_FILE) -> list[RunConfig]:
    if not runs_file.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {runs_file}. "
            f"Make sure you are running the command from the project folder."
        )
    raw = json.loads(runs_file.read_text(), parse_float=Decimal)
    return [RunConfig.model_validate(r) for r in raw]
