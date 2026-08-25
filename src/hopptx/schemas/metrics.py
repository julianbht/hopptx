from decimal import Decimal

from pydantic import BaseModel, ConfigDict


class Metrics(BaseModel):
    model_config = ConfigDict(frozen=True)

    # Sign-up metrics
    first_sign_up_date: str
    last_sign_up_date: str
    total_sign_ups: int
    withdrawn_sign_ups: int
    total_attended_sign_ups: int
    total_evaluation_count: int
    cost_per_training: Decimal
    total_training_cost: Decimal

    # Attendance metrics
    first_attended_date: str
    first_attended_topic: str
    last_attended_date: str
    last_attended_topic: str
    total_sessions: int
    unique_topics: int
    unique_attendees: int

    # Top topics as (name, count) pairs
    top_topics: list[tuple[str, int]]

    # Programs grouped by category
    programs_by_category: dict[str, list[str]]
