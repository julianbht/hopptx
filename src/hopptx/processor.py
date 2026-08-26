import logging
import re
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd
from rapidfuzz import fuzz, process

from hopptx.categories import categorize_topics, load_category_map
from hopptx.schemas.metrics import Metrics
from hopptx.topics import clean_topic_name as _clean_topic_name

log = logging.getLogger(__name__)


def _format_date(date: pd.Timestamp) -> str:
    """Format date as 'DD Month YYYY' (e.g., '24 June 2025')."""
    return f"{date.day} {date.strftime('%B %Y')}"


_COURSE_LEVEL_RE = re.compile(r'\d{3}')


def _differ_only_by_course_level(a: str, b: str) -> bool:
    """Return True if two names are identical except for a course level number (e.g. 101 vs 201)."""
    return _COURSE_LEVEL_RE.sub('', a) == _COURSE_LEVEL_RE.sub('', b) and a != b


def _strip_company_prefix(topics: pd.Series, company: str) -> pd.Series:
    """Strip a leading ``{company} - `` prefix from topic names.

    This merges e.g. "Acme Corp - Coding 101" with "Coding 101" into
    a single topic.  Returns the cleaned series and logs merged groups.
    """
    prefix = f"{company} - "
    stripped = topics.apply(lambda t: t[len(prefix):] if t.startswith(prefix) else t)

    # Log which topics were merged by prefix stripping
    merged: dict[str, set[str]] = {}
    for orig, clean in zip(topics, stripped):
        if orig != clean:
            merged.setdefault(clean, set()).add(orig)
    # Only report groups where a prefixed variant coexists with the bare name
    coexisting = {k: v for k, v in merged.items() if k in topics.values}
    if coexisting:
        log.warning(
            f"Stripped company prefix ({len(coexisting)} group(s)):"
        )
        for clean, originals in sorted(coexisting.items()):
            log.warning(f"  {sorted(originals)} merged into '{clean}'")

    return stripped


def _build_topic_column(
    filtered: pd.DataFrame,
    group_topics: bool,
    clean_topic_names: bool,
    topic_fuzzy_match_threshold: int,
    strip_company_prefix: bool = False,
    company: str | None = None,
) -> pd.Series:
    """Build the derived 'Topic' column from 'Event Type Name'.

    Steps (when enabled):
    0. Strip company prefix (e.g. "Acme - Coding 101" → "Coding 101").
    1. Clean names (strip Q suffixes) — always done when group_topics is True,
       so that fuzzy matching operates on normalised names.
    2. Fuzzy-match remaining near-duplicates (misspellings, word reordering).
    3. If clean_topic_names is True, ensure the final display names are cleaned
       (matters when group_topics is False).
    """
    topics = filtered["Event Type Name"].copy()

    # Step 0 — strip company prefix
    if strip_company_prefix:
        if company is None:
            raise ValueError(
                "company must be provided when strip_company_prefix is True"
            )
        topics = _strip_company_prefix(topics, company)

    # Step 1 — clean for grouping (always when grouping is on)
    if group_topics or clean_topic_names:
        original_names = filtered["Event Type Name"]
        topics = topics.apply(_clean_topic_name)

        # Log which topics were grouped by name cleaning
        cleaned_to_originals: dict[str, set[str]] = {}
        for orig, cleaned in zip(original_names, topics):
            if orig != cleaned:
                if cleaned not in cleaned_to_originals:
                    cleaned_to_originals[cleaned] = set()
                cleaned_to_originals[cleaned].add(orig)
        # Only show groups where multiple originals mapped to the same cleaned name
        grouped = {k: v for k, v in cleaned_to_originals.items() if len(v) > 1}
        if grouped:
            log.info(f"Grouped by name cleaning ({len(grouped)} group(s)):")
            for cleaned, originals in sorted(grouped.items()):
                log.info(f"  {sorted(originals)} -> '{cleaned}'")

    # Step 2 — fuzzy match on the (possibly cleaned) names
    if group_topics:
        unique_topics = sorted(topics.unique())
        fuzzy_matches: list[tuple[str, str, float, str]] = []
        rename_map: dict[str, str] = {}
        matched: set[str] = set()

        for topic in unique_topics:
            if topic in matched:
                continue

            candidates = [t for t in unique_topics if t not in matched and t != topic]
            if not candidates:
                continue

            result = process.extractOne(
                topic, candidates, scorer=fuzz.ratio,
                score_cutoff=topic_fuzzy_match_threshold,
            )

            if result is not None:
                best_result = result
                # Also try token_set_ratio for word reordering
                result_tsr = process.extractOne(
                    topic, candidates, scorer=fuzz.token_set_ratio,
                    score_cutoff=topic_fuzzy_match_threshold,
                )
                if result_tsr is not None:
                    if result_tsr[1] > best_result[1]:
                        best_result = result_tsr
                    elif result_tsr[1] == best_result[1] and result_tsr[0] != best_result[0]:
                        best_result = result_tsr

                match_name, best_score, _ = best_result
                if _differ_only_by_course_level(topic, match_name):
                    log.info(
                        f"  Skipping fuzzy match (course level difference): "
                        f"'{topic}' vs '{match_name}'"
                    )
                    continue
                matched.add(topic)
                matched.add(match_name)

                # Keep the shorter name as canonical (more likely the "clean" one)
                canonical = topic if len(topic) <= len(match_name) else match_name
                rename_map[topic] = canonical
                rename_map[match_name] = canonical
                fuzzy_matches.append((topic, match_name, best_score, canonical))

        if fuzzy_matches:
            topics = topics.replace(rename_map)
            log.warning(
                f"Fuzzy matched {len(fuzzy_matches)} pair(s) "
                f"(threshold: {topic_fuzzy_match_threshold}%):"
            )
            for topic_a, topic_b, score, canonical in fuzzy_matches:
                log.warning(
                    f"  {int(score)}%  '{topic_a}' + '{topic_b}' -> '{canonical}'"
                )

    return topics


def compute_metrics(
    df: pd.DataFrame,
    start: str,
    end: str,
    exclude_canceled: bool,
    cost_per_training: Decimal,
    evaluation_rate_percentage: Decimal,
    top_topics_count: int,
    top_topics_hard_cutoff: int,
    group_topics: bool,
    clean_topic_names: bool,
    topic_fuzzy_match_threshold: int,
    category_reference: str,
    category_fuzzy_match_threshold: int,
    strip_company_prefix: bool = False,
    company: str | None = None,
) -> Metrics:
    """Compute presentation metrics from a report DataFrame."""
    log.info(f"Processing DataFrame with {len(df)} rows")

    df = df.copy()
    df["Start Date & Time"] = pd.to_datetime(df["Start Date & Time"])
    df["Event Created Date & Time"] = pd.to_datetime(df["Event Created Date & Time"])

    start_dt = pd.to_datetime(start)
    end_dt = pd.to_datetime(end) + pd.Timedelta(days=1)  # Make end inclusive

    date_mask = (df["Start Date & Time"] >= start_dt) & (df["Start Date & Time"] < end_dt)
    filtered = df[date_mask]

    if filtered.empty:
        raise ValueError(
            f"No sign-ups found between {start} and {end}. "
            f"Check that the 'start' and 'end' dates in config/runs.json "
            f"match the date range of your data."
        )

    log.info(f"Found {len(filtered)} sign-ups in date range {start} to {end}")

    # Sign-up metrics (from Event Created Date & Time, before filtering canceled)
    all_sign_ups = filtered.copy()

    first_sign_up = all_sign_ups["Event Created Date & Time"].min()
    last_sign_up = all_sign_ups["Event Created Date & Time"].max()
    total_sign_ups = len(all_sign_ups)

    withdrawn_count = all_sign_ups["Canceled"].sum()
    total_attended_sign_ups = total_sign_ups - int(withdrawn_count)
    total_training_cost = Decimal(total_attended_sign_ups) * cost_per_training

    # Attendance metrics (filter out canceled if requested)
    if exclude_canceled:
        canceled_count = filtered["Canceled"].sum()
        canceled_mask = filtered["Canceled"] == True
        filtered = filtered[~canceled_mask]
        if filtered.empty:
            raise ValueError(
                f"All {canceled_count} sign-ups between {start} and {end} were canceled, "
                f"so there is no attendance data to report. "
                f"If you want to include canceled sign-ups, set 'exclude_canceled' "
                f"to false in config/runs.json."
            )
        log.info(f"Excluded {canceled_count} canceled sign-ups")

    # Build derived Topic column (grouping + display cleaning)
    filtered = filtered.copy()
    filtered["Topic"] = _build_topic_column(
        filtered, group_topics, clean_topic_names, topic_fuzzy_match_threshold,
        strip_company_prefix=strip_company_prefix, company=company,
    )

    # Sort by Start Date & Time ascending for attendance
    filtered = filtered.sort_values("Start Date & Time", ascending=True)

    # First attended
    first_row = filtered.iloc[0]
    first_date = first_row["Start Date & Time"]
    first_date_str = _format_date(first_date)
    first_topic = first_row["Topic"]

    # Last attended
    last_row = filtered.iloc[-1]
    last_date = last_row["Start Date & Time"]
    last_date_str = _format_date(last_date)
    last_topic = last_row["Topic"]

    # Session/unique counts
    total_sessions = (
        filtered[["Event Type Name", "Start Date & Time"]]
        .drop_duplicates()
        .shape[0]
    )
    total_evaluation_count = int(
        (
            Decimal(total_sessions)
            * (evaluation_rate_percentage / Decimal("100"))
        ).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    )
    unique_topics = filtered["Topic"].nunique()
    unique_attendees = filtered["Invitee Email"].nunique()

    # Top topics by sign-up count
    topic_counts = filtered["Topic"].value_counts()
    unique_attendance_levels = sorted(topic_counts.unique(), reverse=True)
    if len(unique_attendance_levels) < top_topics_count:
        log.warning(
            f"Only {len(unique_attendance_levels)} distinct attendance level(s) found, "
            f"but top_topics_count is {top_topics_count}. "
            f"The top programs list will show all available levels."
        )

    # Top N ranks (most attended), including all ties at each rank.
    cutoff_index = min(top_topics_count, len(unique_attendance_levels)) - 1
    cutoff_count = int(unique_attendance_levels[cutoff_index])
    top_topic_counts = topic_counts[topic_counts >= cutoff_count]
    top_topic_rows = (
        top_topic_counts.rename_axis("topic")
        .reset_index(name="count")
        .sort_values(["count", "topic"], ascending=[False, True], kind="mergesort")
    )
    if len(top_topic_rows) > top_topics_hard_cutoff:
        log.info(
            f"Top programs hard cutoff applied: keeping "
            f"{top_topics_hard_cutoff} of {len(top_topic_rows)} candidate program(s)."
        )
        top_topic_rows = top_topic_rows.head(top_topics_hard_cutoff)
    top_topics = [
        (str(row.topic), int(row.count)) for row in top_topic_rows.itertuples(index=False)
    ]

    # Categorize attended topics using the reference document
    category_map = load_category_map(category_reference)
    attended_topics = sorted(filtered["Topic"].unique())
    programs_by_category = categorize_topics(
        attended_topics, category_map, category_fuzzy_match_threshold,
    )

    return Metrics(
        # Sign-up metrics
        first_sign_up_date=_format_date(first_sign_up),
        last_sign_up_date=_format_date(last_sign_up),
        total_sign_ups=total_sign_ups,
        withdrawn_sign_ups=withdrawn_count,
        total_attended_sign_ups=total_attended_sign_ups,
        total_evaluation_count=total_evaluation_count,
        cost_per_training=cost_per_training,
        total_training_cost=total_training_cost,

        # Attendance metrics
        first_attended_date=first_date_str,
        first_attended_topic=first_topic,
        last_attended_date=last_date_str,
        last_attended_topic=last_topic,
        total_sessions=total_sessions,
        unique_topics=unique_topics,
        unique_attendees=unique_attendees,

        # Top topics
        top_topics=top_topics,

        # Programs by category
        programs_by_category=programs_by_category,
    )
