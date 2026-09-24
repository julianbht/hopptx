import logging
from pathlib import Path

from docx import Document
from rapidfuzz import fuzz, process

log = logging.getLogger(__name__)


def load_category_map(path: str) -> dict[str, str]:
    """Read the Workshop_descriptions.docx table and return {course_name: category}."""
    doc_path = Path(path)
    if not doc_path.exists():
        raise FileNotFoundError(
            f"Category reference document not found: {path}. "
            f"Please check the 'category_reference' path in config/runs.json."
        )

    doc = Document(str(doc_path))
    category_map: dict[str, str] = {}

    for table in doc.tables:
        for i, row in enumerate(table.rows):
            if i == 0:  # skip header row
                continue
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) < 2:
                raise ValueError(
                    f"Row {i} in category reference has fewer than 2 columns: {cells}"
                )
            course_name, category = cells[0], cells[1]
            if not course_name or not category:
                raise ValueError(
                    f"Row {i} in category reference has empty course name or category: {cells}"
                )
            category_map[course_name] = category

    if not category_map:
        raise ValueError(f"No courses found in category reference: {path}")

    log.info(f"Loaded {len(category_map)} courses from category reference")
    return category_map


def categorize_topics(
    topics: list[str],
    category_map: dict[str, str],
    fuzzy_threshold: int,
) -> dict[str, list[str]]:
    """Match attended topics to categories from the reference document.

    Returns {category: [topic1, topic2, ...]} with topics sorted alphabetically.
    Topics that match no reference course are placed in "Other".
    """
    reference_names = list(category_map.keys())
    result: dict[str, list[str]] = {}
    unmatched: list[str] = []

    for topic in topics:
        # Try exact match first
        if topic in category_map:
            category = category_map[topic]
            result.setdefault(category, []).append(topic)
            continue

        # Fuzzy match against reference course names
        match = process.extractOne(
            topic, reference_names, scorer=fuzz.ratio,
            score_cutoff=fuzzy_threshold,
        )

        # Also try token_set_ratio for word reordering
        match_tsr = process.extractOne(
            topic, reference_names, scorer=fuzz.token_set_ratio,
            score_cutoff=fuzzy_threshold,
        )

        best = None
        if match and match_tsr:
            best = match if match[1] >= match_tsr[1] else match_tsr
        elif match:
            best = match
        elif match_tsr:
            best = match_tsr

        if best is not None:
            matched_name, score, _ = best
            category = category_map[matched_name]
            result.setdefault(category, []).append(topic)
            log.warning(
                f"  Category fuzzy match: {int(score)}%  "
                f"'{topic}' -> '{matched_name}' (category: {category})"
            )
        else:
            unmatched.append(topic)

    if unmatched:
        log.warning(
            f"Could not find category for {len(unmatched)} topic(s), "
            f"placing in 'Other': {unmatched}"
        )
        result.setdefault("Other", []).extend(unmatched)

    # Sort topics within each category
    for category in result:
        result[category].sort()

    return result
