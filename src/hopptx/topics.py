import re


def clean_topic_name(name: str) -> str:
    """Clean topic name by removing quarter suffixes like '- Q1', ' – Q2', etc."""
    # Remove quarter indicators with various dash types and spacing
    # Patterns: " - Q1", "-Q1", " – Q1", "–Q1", " -Q1", "Quarter 1", etc.
    return re.sub(r'[\s\-–—]*(?:Q[1-4]|Quarter [1-4])', '', name, flags=re.IGNORECASE).strip()
